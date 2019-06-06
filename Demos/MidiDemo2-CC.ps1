Import-Module "C:\Repos\Windows-10-PowerShell-MIDI\PeteBrown.PowerShellMidi\bin\Debug\PeteBrown.PowerShellMidi.dll" -Verbose

Get-MidiOutputDeviceInformation

$outputPort = Get-MidiOutputPort -Id "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}"

Send-MidiControlChangeMessage -Port $outputPort -Channel 0 -Controller 0 -Value 64