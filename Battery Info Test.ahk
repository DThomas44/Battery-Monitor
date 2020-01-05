/*
    Battery Info Test
    Author: Daniel Thomas
    Date: 1/5/2019

    NOT ORIGINAL CODE
    Basically a direct pull of https://autohotkey.com/board/topic/7022-acbattery-status/
    with some formatting changes.
*/
;<=====  Script Settings  ====================================================>
#SingleInstance Force
#NoEnv

;<=====  Main  ===============================================================>
VarSetCapacity(powerstatus, 1+1+1+1+4+4)
success := DllCall("kernel32.dll\GetSystemPowerStatus", "uint", &powerstatus)
acLineStatus := ReadInteger(&powerstatus, 0, 1, false)
batteryFlag := ReadInteger(&powerstatus, 1, 1 , false)
batteryLifePercent := ReadInteger(&powerstatus, 2, 1, false)
batteryLifeTime := ReadInteger(&powerstatus, 4, 4, false)
batteryFullLifeTime := ReadInteger(&powerstatus, 8, 4, false)

output := % "AC Status: " . (acLineStatus?"Charging":"Discharging")
    . "`nBattery Flag: " . batteryFlag
    . "`nBattery Life (percent): " . batteryLifePercent
    . "`nBattery Life (time): " . FormatSeconds(batteryLifeTime)
    . "`nBattery Life (full time): " . FormatSeconds(batteryFullLifeTime)
MsgBox, % output
ExitApp

;<=====  Functions  ==========================================================>
ReadInteger(p_address, p_offset, p_size, p_hex=true){
    value = 0
    old_FormatInteger := a_FormatInteger
    if (p_hex){
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
