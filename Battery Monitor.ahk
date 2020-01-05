/*
    Battery Monitor
    Author: Daniel Thomas
    Date: 12/27/2019

    Utilizing:
    https://autohotkey.com/board/topic/7022-acbattery-status/

*/
;<=====  System Settings  =====================================================>
#SingleInstance Force
#Persistent
#NoEnv
SetBatchLines, -1

;<=====  Script Settings  =====================================================>
INIRead, debug, % A_ScriptDir . "\settings.ini", BatteryMonitor, debug, false
INIRead, notifyUser, % A_ScriptDir . "\settings.ini", BatteryMonitor, notifyUser, true
INIRead, useTTS, % A_ScriptDir . "\settings.ini", BatteryMonitor, useTTS, true
INIRead, log, % A_ScriptDir . "\settings.ini", BatteryMonitor, log, true
INIRead, lowBatteryPercent, % A_ScriptDir . "\settings.ini", BatteryMonitor, lowBatteryPercent, 25
INIRead, criticalBatteryPercent, % A_ScriptDir . "\settings.ini", BatteryMonitor, criticalBatteryPercent, 5
INIRead, TTSVolumeOverride, % A_ScriptDir . "\settings.ini", BatteryMonitor, TTSVolumeOverride, 75

;<=====  Timers  ==============================================================>
SetTimer, getStatus, 10000

;<=====  Start TTS  ===========================================================>
tts := ComObjCreate("sapi.SpVoice")

;<=====  Start Logging  =======================================================>
if (log && !LogStart()) {
    MsgBox, % "Unable to start logging battery health.`nPlease contact IT for assistance."
    ExitApp
}
return

;<=====  Labels  ==============================================================>
getStatus:
    ; Get current battery status
    VarSetCapacity(powerstatus, 1+1+1+1+4+4)
    success := DllCall("kernel32.dll\GetSystemPowerStatus", "uint", &powerstatus)
    acLineStatus := ReadInteger(&powerstatus, 0, 1, false)
    batteryFlag := ReadInteger(&powerstatus, 1, 1 , false)
    batteryLifePercent := ReadInteger(&powerstatus, 2, 1, false)
    batteryLifeTime := ReadInteger(&powerstatus, 4, 4, false)
    batteryFullLifeTime := ReadInteger(&powerstatus, 8, 4, false)

    ; Log (un)plugged status
    if (acLineStatus != pACLineStatus) {
        ; Logging
        if log {
            Log("System " . (acLineStatus?"plugged in":"unplugged")
                . ". Battery level: " . batteryLifePercent . "%`n")
        }

        ; Plugged in actions
        if acLineStatus {
            pluginTime := A_Now
            pluginRequested := false
            pluginRequestedCriticalCount := 0
        }

        ; Unplugged actions
        if !acLineStatus {
            unplugTime := A_Now
            notifyFullCharge := false
        }

        if debug {
            MsgBox,,, % "System " . (acLineStatus?"plugged in":"unplugged")
                . "!`nBattery level: " . batteryLifePercent . "%`n"
                . "Estimated battery time: " . FormatSeconds(batteryLifeTime), 30
        }

    }

    ; Notify of full charge
    if ((batteryLifePercent == 100) && acLineStatus && !notifyFullCharge) {
        if log {
            Log("Battery fully charged.")
        }
        if useTTS {
            SoundSet, % TTSVolumeOverride
            tts.speak("Battery full")
        }
        if notifyUser {
            MsgBox, 4144,, % "System fully charged.`n", 30
        }
        notifyFullCharge := true
    }

    ; Notify of low charge
    if ((batteryLifePercent < lowBatteryPercent) && !acLineStatus && !pluginRequested) {
        if log {
            Log("Low battery notification sent to user. Battery level: "
                . batteryLifePercent . "%`n")
        }
        if useTTS {
            SoundSet, % TTSVolumeOverride
            tts.speak("Battery log, please plug in")
        }
        if notifyUser {
            MsgBox, 4144,, % "System charge under " . lowBatteryPercent . "%`nPlease plug in soon.`n"
                . "Estimated battery time: " . FormatSeconds(batteryLifeTime), 30
        }
        pluginRequested := true
    }

    ; Notify of critical charge
    if ((batteryLifePercent < criticalBatteryPercent) && !acLineStatus && !pluginRequestedCritical) {
        pluginRequestedCriticalCount++
        if log {
            Log("Critical battery notification sent to user. Battery level: "
                . batteryLifePercent . "% Notification number: " . pluginRequestedCriticalCount)
        }
        if useTTS {
            SoundSet, % TTSVolumeOverride
            tts.speak("Battery critical, please plug in now.")
        }
        if notifyUser {
            MsgBox, 4112,, % "System charge under " . criticalBatteryPercent . "%`nEstimated battery time remainging: "
                . FormatSeconds(batteryLifeTime) . " (h:mm:ss)`nRequest number: " . pluginRequestedCriticalCount, 5
        }
    }

    ; Update Previous Status variables
    pACLineStatus := acLineStatus
    pBatteryFlag := batteryFlag
    pBatteryLifePercent := batteryLifePercent
    pBatteryLifeTime := batteryLifeTime
    pBatteryFullLifeTime := batteryFullLifeTime
    return

;<=====  Functions  ===========================================================>
FormatSeconds(NumberOfSeconds){
    time := 19990101
    time += NumberOfSeconds, seconds
    FormatTime, mmss, %time%, mm:ss
    return Format("{0:2i}", (NumberOfSeconds//3600)) ":" mmss
}

Log(text){
    file := FileOpen(A_ScriptDir . "\logs\" . A_MM . A_DD . A_YYYY . ".txt", "a")
    FormatTime, logTime, A_Now, MM/dd/yyyy h:mm:sstt
    file.write("[" . logTime . "] " . text)
    file.Close()
}

LogStart(){
    ifNotExist, % A_ScriptDir . "\logs\"
        FileCreateDir, % A_ScriptDir . "\logs"
    file := FileOpen(A_ScriptDir . "\logs\" . A_MM . A_DD . A_YYYY . ".txt", "a")
    FormatTime, logTime, A_Now, MM/dd/yyyy h:mm:sstt
    file.write("[" . logTime . "] Logging started.`n")
    file.Close()
    Return, 1
}

ReadInteger(p_address, p_offset, p_size, p_hex=true){
    value = 0
    old_FormatInteger := a_FormatInteger
    if (p_hex) {
        SetFormat, integer, hex
    }
    else {
        SetFormat, integer, dec
    }
    loop, %p_size%
        value := value+(*((p_address+p_offset)+(a_Index-1)) << (8*(a_Index-1)))
    SetFormat, integer, %old_FormatInteger%
    return, value
}