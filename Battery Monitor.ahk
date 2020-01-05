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
if fileExist(A_ScriptDir . "\settings.ini") {
    INIRead, debug, % A_ScriptDir . "\settings.ini", BatteryMonitor, debug, false
    INIRead, notifyUser, % A_ScriptDir . "\settings.ini", BatteryMonitor, notifyUser, true
    INIRead, useTTS, % A_ScriptDir . "\settings.ini", BatteryMonitor, useTTS, true
    INIRead, log, % A_ScriptDir . "\settings.ini", BatteryMonitor, log, true
} else {
    debug := false
    notifyUser := true
    useTTS := true
    log := true
}

;<=====  Timers  ==============================================================>
SetTimer, getStatus, 2000

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

    /*
    ; Test output
    output := % "AC Status: " . (acLineStatus?"Charging":"Discharging")
        . "`nBattery Flag: " . batteryFlag
        . "`nBattery Life (percent): " . batteryLifePercent
        . "`nBattery Life (time): " . FormatSeconds(batteryLifeTime)
        . "`nBattery Life (full time): " . FormatSeconds(batteryFullLifeTime)
    MsgBox, % output
    */

    ; Do tests & logging
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
            pluginRequestedCritical := false
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
        if notifyUser {
            MsgBox, 4144,, % "System fully charged.`n", 30
        }
        if useTTS {
            SoundSet, 75
            tts.speak("Battery full")
        }
        notifyFullCharge := true
    }

    ; Notify of low charge
    if ((batteryLifePercent < 25) && !acLineStatus && !pluginRequested) {
        if log {
            Log("Low battery notification sent to user. Battery level: "
                . batteryLifePercent . "%`n")
        }
        if notifyUser {
            MsgBox, 4144,, % "System charge under 25%`nPlease plug in soon.`n"
                . "Estimated battery time: " . FormatSeconds(batteryLifeTime), 30
        }
        if useTTS {
            SoundSet, 75
            tts.speak("Battery log, please plug in")
        }
        pluginRequested := true
    }

    ; Notify of critical charge
    if ((batteryLifePercent < 10) && !acLineStatus && !pluginRequestedCritical) {
        if log {
            Log("Critical battery notification sent to user. Battery level: "
                . batteryLifePercent . "%")
        }
        if notifyUser {
            MsgBox, 4112,, % "System charge under 10%`nEstimated battery time remainging: "
                . FormatSeconds(batteryLifeTime) . " (h:mm:ss)", 30
        }
        if useTTS {
            SoundSet, 75
            tts.speak("Battery critical, please plug in now.")
        }
        pluginRequestedCritical := true
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

ReadInteger( p_address, p_offset, p_size, p_hex=true ){
    value = 0
    old_FormatInteger := a_FormatInteger
    if ( p_hex )
      SetFormat, integer, hex
    else
      SetFormat, integer, dec
    loop, %p_size%
      value := value+( *( ( p_address+p_offset )+( a_Index-1 ) ) << ( 8* ( a_Index-1 ) ) )
    SetFormat, integer, %old_FormatInteger%
    return, value
}