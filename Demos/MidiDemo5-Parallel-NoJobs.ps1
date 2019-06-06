Import-Module "C:\Repos\Windows-10-PowerShell-MIDI\PeteBrown.PowerShellMidi\bin\Debug\PeteBrown.PowerShellMidi.dll" -Verbose -Force
Import-Module "C:\TFS\PSDJ\PSDJ" -Force
#Import-Module "C:\Users\JGS\Documents\GitHub\PSLegoEV3\PSLegoEV3WindowsPowerShell" -Force

#Get-MidiOutputDeviceInformation
$outputDeviceIdA = "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}" #loopMIDI Port 1 
$outputDeviceIdB = "\\?\SWD#MMDEVAPI#MIDII_0AB1A389.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}" #loopMIDI Port 2
$global:outputPortA = Get-MidiOutputPort -Id $outputDeviceIdA
$global:outputPortB = Get-MidiOutputPort -Id $outputDeviceIdB

$inputDeviceIdA = "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0004#{504be32c-ccf6-4d2c-b73f-6f8b3747e22b}" #loopMIDI Port 1 
$inputDeviceIdB = "\\?\SWD#MMDEVAPI#MIDII_0AB1A389.P_0004#{504be32c-ccf6-4d2c-b73f-6f8b3747e22b}" #loopMIDI Port 2
$global:inputPortA = Get-MidiInputPort -Id $inputDeviceIdA
$global:inputPortB = Get-MidiInputPort -Id $inputDeviceIdB

Register-ClockEvent
Register-FadeOutEvent
$VerbosePreference = "continue"
break
$VerbosePreference = "Silentlycontinue"


#Load
Invoke-Load -outputPort $global:outputPortA

#Deck A - Reset
Reset-DeckMix  -OutputPort $outputPortA

#reset tempo
Reset-DeckTempo -OutputPort $outputPortA


Start-Deck -outputPort $outputPortA
#master
Set-DeckMaster -outputPort $outputPortA


#Deck B Prep
Select-NextTrack -OutputPort $outputPortB
Invoke-Load -OutputPort $outputPortB 
Select-NextCue -OutputPort $outputPortB 
Enable-Sync -OutputPort $outputPortB 
    

#Deck B  - Reset for Karpus Strong
Set-Volume -Value 127 -outputPort $outputPortB 
Invoke-BassSet -Value 0 -outputPort $outputPortB
Invoke-MidSet -Value 46 -outputPort $outputPortB

Clear-Queue

#Deck B - Karpus Strong
Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown 16 -ScriptBlock {
    Start-Deck  -outputPort $outputPortB
}

Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown 32 -ScriptBlock {
    Invoke-BassFade -BeatDuration 16 -StartValue 0 -EndValue 64 -outputPort $outputPortB 
    Invoke-MidFade -BeatDuration 16 -StartValue 46 -EndValue 64 -outputPort $outputPortB 
    Invoke-BassFade -BeatDuration 16 -StartValue 64 -EndValue 0 -outputPort $outputPortA 
    Invoke-MidFade -BeatDuration 16 -StartValue 64 -EndValue 46 -outputPort $outputPortA 
}

Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown 64 -ScriptBlock {
    Invoke-VolumeFade -BeatDuration 16 -StartValue 127 -EndValue 0 -outputPort $outputPortA 
}
Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown  80 -ScriptBlock {
    Stop-Deck -outputPort $outputPortA
}


break
Invoke-Mix -OutputSource $outputPortA -OutputTarget $outputPortB

Invoke-Mix -OutputSource $outputPortB -OutputTarget $outputPortA


Start-Deck $outputPortA
