#region Procedures
Function start-GUI {
    $ie = new-object -com internetexplorer.application
    $ie.visible = $true
    $IE.Navigate("about:blank")

    return $ie
}

Function Set-GUI {
    param($ie, $content)
    $ie.Document.body.innerHTML = $content
}

Function Start-Deck {
    param(
        $outputPort
    )
    Send-MidiNoteOnMessage -Note 0 -Channel 0 -Velocity 127 -Port $outputPort
}
Function Stop-Deck {
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

Function Invoke-BassSet {
    param(
        $outputPort,
        $Value = 0
    )
    Send-MidiControlChangeMessage -Controller 2 -Value $Value -Channel 0 -Port $outputPort

} #Function VolumeFade 
Function Invoke-MidSet {
    param(
        $outputPort,
        $Value = 0
    )
    Send-MidiControlChangeMessage -Controller 4 -Value $Value -Channel 0 -Port $outputPort

} #Function VolumeFade 


Function Invoke-VolumeFade {
    param(
        $outputPort,
        $StartValue = 0,
        $EndValue = 127,
        $BeatDuration = 8
    )

    Invoke-CCFade -outputPort $outputPort -Channel 0 -Controller 1 -StartValue $StartValue -EndValue $EndValue -BeatDuration $BeatDuration
    
} #Function VolumeFade 

Function Invoke-BassFade {
    param(
        $outputPort,
        $StartValue = 0,
        $EndValue = 64,
        $BeatDuration = 8
    )

    Invoke-CCFade -outputPort $outputPort -Channel 0 -Controller 2 -StartValue $StartValue -EndValue $EndValue -BeatDuration $BeatDuration

} #Function BassFade 

Function Invoke-MidFade {
    param(
        $outputPort,
        $StartValue = 0,
        $EndValue = 64,
        $BeatDuration = 8
    )

    Invoke-CCFade -outputPort $outputPort -Channel 0 -Controller 4 -StartValue $StartValue -EndValue $EndValue -BeatDuration $BeatDuration

} #Function BassFade 

Function Invoke-CCFade {
    param(
        $outputPort,
        $Channel,
        $Controller,
        $StartValue = 0,
        $EndValue = 64,
        $BeatDuration = 8
    )
    $ClockDuration = $Beatduration * 24
    $action = [pscustomobject]@{
        Type       = "CC"
        OutputPort = $outputPort
        Channel    = $Channel
        Controller = $Controller
        Clock      = 1
        ClockEnd   = $ClockDuration
        Start      = $StartValue
        End        = $EndValue
        Increment  = ($EndValue - $StartValue) / $ClockDuration
    }

   
    $global:CurrentActions.Add($action)

} #Function CCFade 

Function Clear-Queue {
    [System.Collections.ArrayList] $global:QueuedActions = @()
    $global:FadeOutReached = $false
}

Function Invoke-QueueAction {
    param(
        [int]$BeatCountDown,
        [boolean] $WaitForFadeOut,
        [ScriptBlock] $ScriptBlock

    )
    $newAction = [pscustomobject]@{ 
        WaitForFadeOut = $WaitForFadeOut
        BeatCountDown  = $BeatCountDown
        ScriptBlock    = $ScriptBlock
    }
    $global:QueuedActions.Add($newAction) | Out-Null
}

Function Register-ClockEvent {

    $global:MidiTiming = 0
    $global:MidiClockQueue = new-object System.Collections.Queue -ArgumentList 20
    
    #Queue for Actions
    [System.Collections.ArrayList] $global:QueuedActions = @()
    [System.Collections.ArrayList] $global:CurrentActions = @()
    
    #check if event is already registered and remove.
    if (Get-EventSubscriber -SourceIdentifier "MidiClockMessageReceived" -ErrorAction SilentlyContinue) {
        UnRegister-Event -SourceIdentifier "MidiClockMessageReceived"  -ErrorAction silentlyContinue
    }

    $global:BeatCount = 0
    $ActionClock = { 
        $global:MidiTiming += 1
        $oneBeat = 24

        $CurrentActionsToRemove = @()
        #Handle current actions
        foreach ($Action in $global:CurrentActions) {
            $Action.Clock++;

            switch ($Action.Type) {
                "CC" {
                    $Value = $Action.Start + [int]($Action.Increment * $Action.Clock)
                    Send-MidiControlChangeMessage -Controller $Action.Controller -Value $Value -Channel $Action.Channel -Port $Action.outputPort
                }
            }

            if ($Action.ClockEnd -eq $Action.Clock) {
                #Remove Action from Queue
                $CurrentActionsToRemove += $Action
            
            }
        } #foreach current action

        foreach ($Action in $CurrentActionsToRemove) {
            write-host "Removing $($Action)"
            $global:CurrentActions.Remove($Action)
        }

        if ($global:MidiTiming % $oneBeat -eq 0) {
            #on Beat
            $global:BeatCount++

            if ($global:FadeOutReached -eq $true) {
                $global:FadeOutReached = $false
                $global:QueuedActions | Where-Object WaitForFadeOut -eq $True | Foreach-Object { $_.WaitForFadeOut = $false }
            }

            #Actions
            $ActionsToRemove = @()
            Foreach ($Action in ($global:QueuedActions | Where-Object WaitForFadeOut -eq $false )) {
                write-host $Action.BeatCountDown
                $Action.BeatCountDown--

                if ($Action.BeatCountDown -in 0, -1) {
                    #Execute Action block
                    write-host "Executing $($Action.ScriptBlock)"
                    & $Action.ScriptBlock + { $Value }

                    #Remove Action from Queue
                    $ActionsToRemove += $Action
                }
            }
            foreach ($Action in $ActionsToRemove) {
                write-host "Removing $($Action.ScriptBlock)"
                $global:QueuedActions.Remove($Action)
            }

            if ($global:RobotEnabled -eq $true) {
                if ($global:BeatCount % 2 -eq 0) {
                    Invoke-EV3Turn -Direction Right -Steps 70 
                }
                else {
                    Invoke-EV3Turn -Direction Left -Steps 70 
                }
            }
            # BPM
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
    Register-ObjectEvent -InputObject $inputPortA -EventName MidiClockMessageReceived -SourceIdentifier "MidiClockMessageReceived" -Action $ActionClock  -Verbose
}
Function Invoke-FadeOutReached {
    write-verbose "Fade Out Reached"
    $global:FadeOutReached = $true
}

Function Register-FadeOutEvent {
    #check if events are already registered and remove.
    if (Get-EventSubscriber -SourceIdentifier "NoteOnMessageReceivedA" -ErrorAction SilentlyContinue) {
        UnRegister-Event -SourceIdentifier "NoteOnMessageReceivedA"  -ErrorAction silentlyContinue
    }
    if (Get-EventSubscriber -SourceIdentifier "NoteOnMessageReceivedB" -ErrorAction SilentlyContinue) {
        UnRegister-Event -SourceIdentifier "NoteOnMessageReceivedB"  -ErrorAction silentlyContinue
    }
    $ActionNoteOn = {
        $Note = $event.SourceArgs.Note
        switch ($Note) {
            26 { 
                #26 = D1 = Fadeout reached
                Invoke-FadeOutReached
            }
        }
    }

    UnRegister-Event -SourceIdentifier "NoteOnMessageReceivedA"  -ErrorAction SilentlyContinue
    UnRegister-Event -SourceIdentifier "NoteOnMessageReceivedB"  -ErrorAction SilentlyContinue
    Register-ObjectEvent -InputObject $inputPortA -EventName NoteOnMessageReceived -SourceIdentifier "NoteOnMessageReceivedA" -Action $ActionNoteOn  -Verbose
    Register-ObjectEvent -InputObject $inputPortB -EventName NoteOnMessageReceived -SourceIdentifier "NoteOnMessageReceivedB" -Action $ActionNoteOn  -Verbose
}
Function Invoke-RemixDeckSample {
    param($slot, $cell, $outputPort)

    switch ($slot) {
        1 { $slotCCDigit = 5 }
        2 { $slotCCDigit = 6 }
        
    }
    switch ($cell) {
        1 { $cellCCDigit = 0 }
        2 { $cellCCDigit = 1 }
        
    }
    Send-MidiNoteOnMessage  -Note "$slotCCDigit$cellCCDigit" -Channel 0 -Velocity 127 -Port $outputPort
    Start-Sleep -Milliseconds 50
    Send-MidiNoteOffMessage  -Note "$slotCCDigit$cellCCDigit" -Channel 0 -Velocity 127 -Port $outputPort
}

Function Select-NextTrack($OutputPort) {
    # Next
    Send-MidiNoteOnMessage -Note 4 -Channel 0 -Velocity 127 -Port $OutputPort
    Start-Sleep -Milliseconds 100
}
Function Select-NextCue($OutputPort) {
    #Next que
    Send-MidiNoteOnMessage -Note 6 -Channel 0 -Velocity 127 -Port $OutputPort
    Start-Sleep -Milliseconds 100
}
Function Invoke-Load($OutputPort) {
    #Load
    Send-MidiNoteOnMessage -Note 3 -Channel 0 -Velocity 127 -Port $OutputPort
    Start-Sleep -Milliseconds 100
}
Function Enable-Sync($OutputPort) {
    #Sync
    Send-MidiNoteOnMessage -Note 2 -Channel 0 -Velocity 127 -Port $OutputPort
    Start-Sleep -Milliseconds 100
}
Function Invoke-Mix {    
    param(
        $OutputSource, 
        $OutputTarget,
        $BeatCountTargetStart = 16,
        $BeatCountTargetEQFadeIn = 32,
        $BeatCountSourceFadeOut = 64,
        $BeatCountSourceStop = 80
        
        )
    #Deck Target load
    # Next
    Select-NextTrack -OutputPort $OutputTarget
    Invoke-Load -OutputPort $OutputTarget

    Start-Sleep -Milliseconds 500
    #Next que
    Select-NextCue -OutputPort $OutputTarget

    Enable-Sync -OutputPort  $OutputTarget

    #Deck B  - Reset for Karpus Strong
    Set-Volume -Value 127 -outputPort $OutputTarget 
    Invoke-BassSet -Value 0 -outputPort $OutputTarget
    Invoke-MidSet -Value 46 -outputPort $OutputTarget

    Clear-Queue
    
    Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown $BeatCountTargetStart -ScriptBlock {
        Start-Deck -outputPort $OutputTarget
        Start-Sleep -Milliseconds 1000
        Enable-Sync -OutputPort  $OutputTarget
    }

    Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown $BeatCountTargetEQFadeIn -ScriptBlock {
        Invoke-BassFade -BeatDuration 16 -StartValue 0 -EndValue 64 -outputPort $OutputTarget 
        Invoke-MidFade -BeatDuration 16 -StartValue 46 -EndValue 64 -outputPort $OutputTarget 
        Invoke-BassFade -BeatDuration 16 -StartValue 64 -EndValue 0 -outputPort $OutputSource 
        Invoke-MidFade -BeatDuration 16 -StartValue 64 -EndValue 46 -outputPort $OutputSource 
    }
    Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown $BeatCountSourceFadeOut -ScriptBlock {
        Invoke-VolumeFade -BeatDuration 16 -StartValue 127 -EndValue 0 -outputPort $OutputSource 
    }
    Invoke-QueueAction -WaitForFadeOut $true -BeatCountDown $BeatCountSourceStop -ScriptBlock {
        Stop-Deck -outputPort $OutputSource
    }
}

Function Reset-DeckMix ($OutputPort) {
    Set-Volume -Value 127 -outputPort $OutputPort
    Invoke-BassSet -Value 64 -outputPort $OutputPort
    Invoke-MidSet -Value 64 -outputPort $OutputPort
}

Function Reset-DeckTempo ($OutputPort){
    #reset tempo
    Send-MidiControlChangeMessage -Controller 3  -Value 64 -Channel 0 -Port $OutputPort
}

Function Set-DeckMaster($OutputPort){
    #master
    Send-MidiNoteOnMessage -Note 1 -Channel 0 -Velocity 127 -Port $OutputPort
}
#endregion Procedures
