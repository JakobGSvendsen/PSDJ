Import-Module "C:\Repos\Windows-10-PowerShell-MIDI\PeteBrown.PowerShellMidi\bin\Debug\PeteBrown.PowerShellMidi.dll" -Verbose
Get-MidiOutputDeviceInformation

$outputDeviceIdA = "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}" #loopMIDI Port 1 
$outputDeviceIdB = "\\?\SWD#MMDEVAPI#MIDII_0AB1A389.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}" #loopMIDI Port 2
$global:outputPortA = Get-MidiOutputPort -Id $outputDeviceIdA
$global:outputPortB = Get-MidiOutputPort -Id $outputDeviceIdB

$inputDeviceIdA = "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0004#{504be32c-ccf6-4d2c-b73f-6f8b3747e22b}" #loopMIDI Port 1 
$inputDeviceIdB = "\\?\SWD#MMDEVAPI#MIDII_0AB1A389.P_0004#{504be32c-ccf6-4d2c-b73f-6f8b3747e22b}" #loopMIDI Port 2
$global:inputPortA = Get-MidiInputPort -Id $inputDeviceIdA
$global:inputPortB = Get-MidiInputPort -Id $inputDeviceIdB

UnRegister-Event -SourceIdentifier "MidiClockMessageReceived"  -ErrorAction silentlyContinue

$global:BeatCount = 0
$global:MidiTiming = 0
$global:MidiClockQueue = new-object System.Collections.Queue -ArgumentList 20

$ActionClock = { 
    $global:MidiTiming += 1
    $oneBeat = 24

    if ($global:MidiTiming % $oneBeat -eq 0) {
        #on Beat
        $global:BeatCount++

        if ($global:RobotEnabled -eq $true) {
            if ($global:BeatCount % 2 -eq 0) {
                Invoke-EV3Turn -Direction Right -Steps 70 
            }
            else {
                Invoke-EV3Turn -Direction Left -Steps 70 
            }
        }

        # BPM V2
        $NumberOfBeats = 4
        $global:MidiClockQueue.Enqueue($event.SourceArgs[1].TimeStamp)
        if ($global:MidiClockQueue.Count -eq ($NumberOfBeats + 1)) {
            $global:MidiClockQueue.Dequeue()
            $BeatsTimeSpan = $null
            $Array = $global:MidiClockQueue.ToArray()
            for ($i = 1; $i -lt $array.count; $i++) {
                $BeatsTimeSpan += ($array[$i] - $array[$i - 1])
            }
            $global:BPM = 60 / ($BeatsTimeSpan.TotalSeconds / ($NumberOfBeats - 1))
            write-verbose "Average BPM from last $NumberOfBeats beats: $global:BPM"
        }       
    } #on Beat

}
$VerbosePreference = "Continue"
$JobClock = Register-ObjectEvent -InputObject $inputPortA -EventName MidiClockMessageReceived -SourceIdentifier "MidiClockMessageReceived" -Action $ActionClock  -Verbose


$global:RobotEnabled = $false