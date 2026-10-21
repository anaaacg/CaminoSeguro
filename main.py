"""CaminoSeguro IA: clasificador de imágenes de Teachable Machine y voz Windows."""
import os
os.environ.setdefault('TF_CPP_MIN_LOG_LEVEL', '2')
import sys
import time
import base64
import queue
import threading
import subprocess
from pathlib import Path

import cv2
import numpy as np
from PySide6.QtCore import Qt, QTimer, Signal, QObject
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QCheckBox, QSpinBox, QProgressBar,
    QListWidget, QScrollArea, QMessageBox, QFrame, QSplitter,
    QSizePolicy,
)

BASE = Path(__file__).resolve().parent
MODEL_DIR = BASE / 'modelo'
MESSAGES = {
    'LIBRE': 'Camino aparentemente libre.',
    'OBSTACULO': 'Precaución, posible obstáculo adelante.',
    'DESNIVEL': 'Atención, posible desnivel.',
    'ESCALERA_BAJA': 'Atención, posible escalera hacia abajo.',
    'ESCALERA_SUBE': 'Atención, posible escalera hacia arriba.',
    'CALLE': 'Precaución, posible salida a la calle.',
}
ORDER = list(MESSAGES)
INFER_GAP = 0.10           # mínimo entre inicios de inferencia; sin hilos acumulados
SAME_ALERT_GAP = 5.0      # no repetir inmediatamente la misma alerta
DIFFERENT_ALERT_GAP = 0.25


def read_labels(path):
    out = []
    for raw in path.read_text(encoding='utf-8-sig').splitlines():
        raw = raw.strip()
        if raw:
            pieces = raw.split(' ', 1)
            out.append((pieces[1] if len(pieces) == 2 and pieces[0].isdigit() else raw).strip().upper())
    return out


class Signals(QObject):
    prediction = Signal(object)
    loaded = Signal(object, object)
    error = Signal(str)
    voice_status = Signal(str)


class WindowsVoice:
    """Un único proceso SAPI persistente. Las frases nuevas cancelan las anteriores."""
    def __init__(self, signals):
        self.signals = signals
        self.q = queue.Queue(maxsize=1)
        self.process = None
        self.active = True
        threading.Thread(target=self._run, daemon=True).start()

    def say(self, phrase):
        try:
            while True:
                self.q.get_nowait()
        except queue.Empty:
            pass
        try:
            self.q.put_nowait(phrase)
        except queue.Full:
            pass

    def _run(self):
        if os.name != 'nt':
            self.signals.voice_status.emit('La voz SAPI requiere Windows')
            return
        code = r'''
$ErrorActionPreference = 'Stop'
try {
    $v = New-Object -ComObject SAPI.SpVoice
    $all = $v.GetVoices()
    for ($i=0; $i -lt $all.Count; $i++) {
        $candidate = $all.Item($i)
        $lang = $candidate.GetAttribute('Language')
        $desc = $candidate.GetDescription()
        if ($lang -match '(^|;)0?C0A|(^|;)040A|(^|;)080A|(^|;)2C0A|(^|;)340A|(^|;)240A|(^|;)3C0A|(^|;)140A' -or $desc -match 'Spanish|Espa|Sabina|Helena|Laura|Pablo') {
            $v.Voice = $candidate
            break
        }
    }
    $v.Rate = 2
    $v.Volume = 100
    [Console]::WriteLine('READY')
    while ($null -ne ($line = [Console]::ReadLine())) {
        if ($line -eq 'EXIT') { break }
        try {
            $phrase = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line))
            # SpeakFlags=3: asíncrono y purga de la frase anterior
            [void]$v.Speak($phrase, 3)
        } catch {}
    }
} catch {
    [Console]::WriteLine('ERROR: ' + $_.Exception.Message)
}
'''
        cmd = base64.b64encode(code.encode('utf-16-le')).decode('ascii')
        try:
            self.process = subprocess.Popen(
                ['powershell.exe', '-NoProfile', '-NonInteractive', '-EncodedCommand', cmd],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                text=True, encoding='utf-8', bufsize=1,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
            ready = self.process.stdout.readline().strip()
            if ready != 'READY':
                raise RuntimeError(ready or 'No se pudo abrir SAPI')
            self.signals.voice_status.emit('Voz lista (Windows SAPI)')
            while self.active:
                try:
                    message = self.q.get(timeout=0.2)
                except queue.Empty:
                    continue
                try:
                    while True:
                        message = self.q.get_nowait()
                except queue.Empty:
                    pass
                encoded = base64.b64encode(message.encode('utf-8')).decode('ascii')
                self.process.stdin.write(encoded + '\n')
                self.process.stdin.flush()
        except Exception as exc:
            self.signals.voice_status.emit('Error de voz: ' + str(exc)[:120])

    def close(self):
        self.active = False
        if self.process is not None:
            try:
                self.process.terminate()
            except Exception:
                pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('CaminoSeguro IA · Escritorio')
        self.resize(1280, 820)
        self.setMinimumSize(840, 550)
        self.model = None
        self.labels = []
        self.cap = None
        self.running = False
        self.loading = False
        self.busy = False
        self.lock = threading.Lock()
        self.last_infer = 0
        self.session = 0
        self.candidate = None
        self.streak = 0
        self.last_spoken_label = None
        self.last_spoken_at = 0
        self.last_any_at = 0
        self.signals = Signals()
        self.signals.prediction.connect(self.show_prediction)
        self.signals.loaded.connect(self.model_ready)
        self.signals.error.connect(self.report_error)
        self.signals.voice_status.connect(self.voice_state)
        self.make_ui()
        self.voice = WindowsVoice(self.signals)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.frame_tick)

    def make_ui(self):
        root = QWidget()
        root.setObjectName('root')
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(20, 15, 20, 15)
        outer.setSpacing(9)
        title = QLabel('✦  CaminoSeguro IA')
        title.setObjectName('title')
        outer.addWidget(title)
        sub = QLabel('Asistente visual experimental · Detección mediante el modelo original')
        sub.setObjectName('muted')
        outer.addWidget(sub)
        controls = QHBoxLayout()
        self.btn_start = QPushButton('▶  Iniciar cámara')
        self.btn_start.clicked.connect(self.start)
        controls.addWidget(self.btn_start)
        self.btn_stop = QPushButton('■  Detener')
        self.btn_stop.clicked.connect(self.stop)
        controls.addWidget(self.btn_stop)
        outer.addLayout(controls)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        outer.addWidget(splitter, 1)
        left = QWidget()
        vl = QVBoxLayout(left)
        vl.setContentsMargins(0, 0, 5, 0)
        self.preview = QLabel('Cámara detenida\nPulsa «Iniciar cámara»')
        self.preview.setObjectName('preview')
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumSize(300, 220)
        self.preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        vl.addWidget(self.preview, 1)
        self.status = QLabel('Modelo listo para cargar desde la carpeta modelo')
        self.status.setObjectName('muted')
        self.status.setWordWrap(True)
        vl.addWidget(self.status)
        splitter.addWidget(left)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        right_panel = QWidget()
        right = QVBoxLayout(right_panel)
        right.setContentsMargins(8, 0, 8, 8)
        right.setSpacing(10)
        label = QLabel('RESULTADO ACTUAL')
        label.setObjectName('muted')
        right.addWidget(label)
        self.current = QLabel('—')
        self.current.setObjectName('detected')
        self.current.setWordWrap(True)
        right.addWidget(self.current)
        self.confidence = QLabel('Confianza: —')
        right.addWidget(self.confidence)
        self.bars = {}
        for category in ORDER:
            caption = QLabel(category.replace('_', ' '))
            bar = QProgressBar()
            bar.setRange(0, 1000)
            bar.setTextVisible(False)
            bar.setFixedHeight(12)
            right.addWidget(caption)
            right.addWidget(bar)
            self.bars[category] = (caption, bar)
        self.voice_on = QCheckBox('Activar alertas de voz')
        self.voice_on.setChecked(True)
        right.addWidget(self.voice_on)
        vrow = QHBoxLayout()
        self.voice_label = QLabel('Iniciando voz...')
        self.voice_label.setObjectName('muted')
        self.voice_label.setWordWrap(True)
        vrow.addWidget(self.voice_label, 1)
        btn_test = QPushButton('🔊 Probar voz')
        btn_test.setObjectName('secondary')
        btn_test.clicked.connect(lambda: self.voice.say('Atención. Posible escalera hacia arriba.'))
        vrow.addWidget(btn_test)
        right.addLayout(vrow)
        settings = QHBoxLayout()
        settings.addWidget(QLabel('Confianza mínima'))
        self.threshold = QSpinBox()
        self.threshold.setRange(1, 100)
        self.threshold.setSuffix(' %')
        self.threshold.setValue(82)
        settings.addWidget(self.threshold)
        settings.addWidget(QLabel('Lecturas'))
        self.readings = QSpinBox()
        self.readings.setRange(1, 10)
        self.readings.setValue(3)
        settings.addWidget(self.readings)
        right.addLayout(settings)
        right.addWidget(QLabel('Historial de alertas'))
        self.history = QListWidget()
        self.history.setMinimumHeight(130)
        right.addWidget(self.history)
        scroll.setWidget(right_panel)
        splitter.addWidget(scroll)
        splitter.setSizes([820, 450])
        footer = QLabel('Prototipo académico: la clasificación no garantiza identificar peligros reales. No usar como única ayuda para caminar.')
        footer.setObjectName('muted')
        footer.setWordWrap(True)
        outer.addWidget(footer)
        self.setStyleSheet('''
            QWidget#root {background:#f3f8f5;color:#173c31;font-size:13px;font-family:Segoe UI,Arial;}
            QLabel#title {font-size:28px;font-weight:750;color:#176956;}
            QLabel#muted {color:#647d73;}
            QLabel#detected {font-size:27px;font-weight:750;color:#176956;}
            QLabel#preview {background:#162a27;color:#bbd0c8;border-radius:14px;font-size:18px;}
            QPushButton {background:#17785f;color:white;border:none;border-radius:8px;padding:10px;font-weight:600;}
            QPushButton:hover {background:#115943;}
            QPushButton#secondary {background:#dcece5;color:#176956;}
            QSpinBox,QListWidget {background:white;border:1px solid #cfddd5;border-radius:7px;padding:5px;}
            QProgressBar {background:#e5efea;border:none;border-radius:4px;}
            QProgressBar::chunk {background:#25a17e;border-radius:4px;}
            QScrollArea {background:transparent;border:none;}
        ''')

    def start(self):
        if self.running or self.loading:
            return
        if self.model is None:
            self.loading = True
            self.btn_start.setEnabled(False)
            self.status.setText('Cargando el modelo original...')
            threading.Thread(target=self.load_model, daemon=True).start()
        else:
            self.open_camera()

    def load_model(self):
        try:
            mp = MODEL_DIR / 'keras_model.h5'
            lp = MODEL_DIR / 'labels.txt'
            if not mp.exists() or not lp.exists():
                raise FileNotFoundError('Coloca keras_model.h5 y labels.txt en la carpeta modelo')
            import tensorflow as tf
            model = tf.keras.models.load_model(str(mp), compile=False)
            labels = read_labels(lp)
            if labels != ORDER:
                raise ValueError('Las seis etiquetas no coinciden con el modelo original o están en otro orden: ' + str(labels))
            if model.output_shape[-1] != len(labels):
                raise ValueError('Cantidad de salidas distinta de cantidad de etiquetas')
            model.predict(np.zeros((1, 224, 224, 3), dtype=np.float32), verbose=0)
            self.signals.loaded.emit(model, labels)
        except Exception as e:
            self.signals.error.emit('Error al cargar modelo: ' + str(e))

    def model_ready(self, model, labels):
        self.model = model
        self.labels = labels
        self.loading = False
        self.btn_start.setEnabled(True)
        self.open_camera()

    def open_camera(self):
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW if os.name == 'nt' else cv2.CAP_ANY)
        if not cap.isOpened():
            cap.release()
            QMessageBox.warning(self, 'Cámara', 'No se pudo abrir la cámara.')
            return
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        self.cap = cap
        self.running = True
        self.session += 1
        self.candidate = None
        self.streak = 0
        self.last_spoken_label = None
        self.last_spoken_at = 0
        self.last_any_at = 0
        self.last_infer = 0
        self.btn_start.setEnabled(False)
        self.status.setText('Cámara activa · Modelo original, RGB 224 × 224')
        self.timer.start(30)

    def frame_tick(self):
        if not self.running or self.cap is None:
            return
        ok, frame = self.cap.read()
        if not ok:
            return
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w = rgb.shape[:2]
        image = QImage(rgb.data, w, h, rgb.strides[0], QImage.Format_RGB888).copy()
        self.preview.setPixmap(QPixmap.fromImage(image).scaled(
            self.preview.size(), Qt.KeepAspectRatio, Qt.FastTransformation))
        now = time.monotonic()
        with self.lock:
            if self.busy or now - self.last_infer < INFER_GAP:
                return
            self.busy = True
            self.last_infer = now
        threading.Thread(target=self.predict, args=(frame.copy(), self.session), daemon=True).start()

    def predict(self, frame, session):
        try:
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            h, w = rgb.shape[:2]
            side = min(h, w)
            # Mismo recorte centrado y normalización de Teachable Machine.
            rgb = rgb[(h-side)//2:(h+side)//2, (w-side)//2:(w+side)//2]
            rgb = cv2.resize(rgb, (224, 224), interpolation=cv2.INTER_AREA)
            data = (rgb.astype(np.float32) / 127.5) - 1.0
            scores = self.model.predict(data[np.newaxis, ...], verbose=0)[0]
            self.signals.prediction.emit((session, list(zip(self.labels, map(float, scores)))))
        except Exception as exc:
            self.signals.error.emit('Error al predecir: ' + str(exc))
        finally:
            with self.lock:
                self.busy = False

    def show_prediction(self, payload):
        session, predictions = payload
        if not self.running or session != self.session:
            return
        name, score = max(predictions, key=lambda row: row[1])
        self.current.setText(name.replace('_', ' '))
        self.confidence.setText(f'Confianza: {score * 100:.1f} %')
        for label, value in predictions:
            if label in self.bars:
                caption, bar = self.bars[label]
                caption.setText(f'{label.replace("_", " ")}: {value * 100:.0f} %')
                bar.setValue(max(0, min(1000, round(value * 1000))))
        if score < self.threshold.value() / 100:
            self.candidate, self.streak = None, 0
            return
        if name == self.candidate:
            self.streak += 1
        else:
            self.candidate, self.streak = name, 1
        if self.streak < self.readings.value():
            return
        if not self.voice_on.isChecked():
            return
        now = time.monotonic()
        if name == self.last_spoken_label and now - self.last_spoken_at < SAME_ALERT_GAP:
            return
        if name != self.last_spoken_label and now - self.last_any_at < DIFFERENT_ALERT_GAP:
            return
        message = MESSAGES[name]
        self.last_spoken_label = name
        self.last_spoken_at = now
        self.last_any_at = now
        self.history.insertItem(0, time.strftime('%H:%M:%S') + ' — ' + message)
        if self.history.count() > 70:
            self.history.takeItem(self.history.count() - 1)
        self.voice.say(message)

    def voice_state(self, message):
        self.voice_label.setText(message)

    def report_error(self, message):
        self.loading = False
        self.btn_start.setEnabled(True)
        self.status.setText(message)
        QMessageBox.warning(self, 'CaminoSeguro IA', message)

    def stop(self):
        if hasattr(self, 'timer'):
            self.timer.stop()
        self.running = False
        self.session += 1
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        self.candidate = None
        self.streak = 0
        if hasattr(self, 'voice'):
            self.voice.say('')  # interrumpe cualquier frase activa
        self.btn_start.setEnabled(not self.loading)
        self.preview.clear()
        self.preview.setText('Cámara detenida\nPulsa «Iniciar cámara»')
        self.status.setText('Detenido')

    def closeEvent(self, event):
        self.stop()
        self.voice.close()
        super().closeEvent(event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
