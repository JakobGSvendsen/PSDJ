Import-Module -Name PoshRSJob
Import-Module "C:\Repos\Windows-10-PowerShell-MIDI\PeteBrown.PowerShellMidi\bin\Debug\PeteBrown.PowerShellMidi.dll" -Verbose -Force


Function Play {
    param(
        $outputPort
    )
    Send-MidiNoteOnMessage -Note 0 -Channel 0 -Velocity 127 -Port $outputPort
}
Function Stop {
    param(
        $outputPort
    )
    Send-MidiNoteOffMessage -Note 0 -Channel 0 -Velocity 127 -Port $outputPort
}
Function Set-Volume {
    param(
        $outputPort,
        $Value = 0
    )
    Send-MidiControlChangeMessage -Controller 1 -Value $Value -Channel 0 -Port $outputPort

} #Function VolumeFade 
Function VolumeFade {
    param(
        $outputPort,
        $StartValue = 0,
        $EndValue = 127,
        $IntervalMiliseconds = 10
    )
    start-rsjob -FunctionsToImport Set-Volume -ModulesToImport "C:\Repos\Windows-10-PowerShell-MIDI\PeteBrown.PowerShellMidi\bin\Debug\PeteBrown.PowerShellMidi.dll" -Verbose:$false -scriptBlock { 
        $using:StartValue..$using:EndValue |
        ForEach {
            Set-Volume -outputPort $using:outputPort -Value $_ 
            Start-Sleep -Milliseconds 10
        }
    } 
        

} #Function VolumeFade 

Function QueueClear {
    [System.Collections.ArrayList] $global:QueuedActions = @()
}
    
Function QueueAction {
    param(
        [int]$BeatCountDown,
        [boolean] $WaitForFadeOut,
        [ScriptBlock] $ScriptBlock
    
    )
    $newAction = @{ 
        WaitForFadeOut = $WaitForFadeOut
        BeatCountDown  = $BeatCountDown
        ScriptBlock    = $ScriptBlock
    }
    $global:QueuedActions.Add($newAction) | Out-Null
}

Function Invoke-FadeOutReached {
    write-verbose "Fade Out Reached"
    $global:FadeOutReached = $true
}

#endregion Procedures
#region Beat Sync

#Get-MidiOutputDeviceInformation
$outputDeviceIdA = "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}" #loopMIDI Port 1 
$outputDeviceIdB = "\\?\SWD#MMDEVAPI#MIDII_0AB1A389.P_0000#{6dc23320-ab33-4ce4-80d4-bbb3ebbf2814}" #loopMIDI Port 2
$global:outputPortA = Get-MidiOutputPort -Id $outputDeviceIdA
$global:outputPortB = Get-MidiOutputPort -Id $outputDeviceIdB

$inputDeviceIdA = "\\?\SWD#MMDEVAPI#MIDII_0AB1A388.P_0004#{504be32c-ccf6-4d2c-b73f-6f8b3747e22b}" #loopMIDI Port 1 
$inputDeviceIdB = "\\?\SWD#MMDEVAPI#MIDII_0AB1A389.P_0004#{504be32c-ccf6-4d2c-b73f-6f8b3747e22b}" #loopMIDI Port 2
$global:inputPortA = Get-MidiInputPort -Id $inputDeviceIdA
$global:inputPortB = Get-MidiInputPort -Id $inputDeviceIdB

UnRegister-Event -SourceIdentifier "MidiClockMessageReceived"  -ErrorAction silentlyContinue

#Clock
$global:BeatCount = 0
$global:MidiTiming = 0
$global:MidiClockQueue = new-object System.Collections.Queue -ArgumentList 20

#Queue for Actions
[System.Collections.ArrayList] $global:QueuedActions = @()

UnRegister-Event -SourceIdentifier "NoteOnMessageReceived"  -ErrorAction SilentlyContinue
$ActionNoteOn = {
    $Note = $event.SourceArgs.Note
    switch ($Note) {
        26 {
            #26 = D1 = Fadeout reached
            Invoke-FadeOutReached
        }
    }
}
$JobFadeOutReached = Register-ObjectEvent -InputObject $inputPortA -EventName NoteOnMessageReceived -SourceIdentifier "NoteOnMessageReceived" -Action $ActionNoteOn  -Verbose


$ActionClock = { 
    $global:MidiTiming += 1
    $oneBeat = 24

    if ($global:MidiTiming % $oneBeat -eq 0) {
        #on Beat
        $global:BeatCount++

        #Actions
        if ($global:FadeOutReached -eq $true) {
            $global:FadeOutReached = $false
            $global:QueuedActions | Where-Object WaitForFadeOut -eq $True | % { $_.WaitForFadeOut = $false }
        }

        #Actions
        $global:QueuedActions
        $ActionsToRemove = @()
        Foreach ($Action in ($global:QueuedActions | Where-Object WaitForFadeOut -eq $false )) {
            write-host $Action.BeatCountDown
            $Action.BeatCountDown--
            if ($Action.BeatCountDown -eq 0) {
                #Execute Action block
                write-host "Executing $($Action.ScriptBlock)"
                & $Action.ScriptBlock

                #Remove Action from Queue
                $ActionsToRemove += $Action
            }
        }

        foreach ($Action in $ActionsToRemove) {
            write-host "Removing $($Action.ScriptBlock)"
            $global:QueuedActions.Remove($Action)
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

break
$VerbosePreference = "Silentlycontinue"
$VerbosePreference = "continue"
Get-EventSubscriber

#Demo
QueueClear
Stop -outputPort $global:outputPortB
Play -outputPort $global:outputPortA

QueueAction -BeatCountDown 4 -ScriptBlock {VolumeFade -outputPort $global:outputPorta }
QueueAction -BeatCountDown 4 -ScriptBlock {VolumeFade -outputPort $global:outputPortB }

