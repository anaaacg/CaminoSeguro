# Voz nativa de Windows: requiere Windows PowerShell 5.1 y System.Speech.
$ErrorActionPreference = 'Stop'
try {
  Add-Type -AssemblyName System.Speech
  $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
  $synth.Rate = -1
  $synth.Volume = 100
  $all = @($synth.GetInstalledVoices() | Where-Object { $_.Enabled })
  $spanish = $all | Where-Object { $_.VoiceInfo.Culture.Name -like 'es-*' } | Select-Object -First 1
  if ($null -ne $spanish) { $synth.SelectVoice($spanish.VoiceInfo.Name) }
  $text = [Console]::In.ReadToEnd().Trim()
  if ($text.Length -gt 0) { $synth.Speak($text) }
  [Console]::Out.WriteLine($synth.Voice.Name + ' (' + $synth.Voice.Culture.Name + ')')
  $synth.Dispose()
  exit 0
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
