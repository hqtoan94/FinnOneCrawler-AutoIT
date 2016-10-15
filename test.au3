$state = "TP Há»“ ChÃ­ Minh"

Func Asc2Unicode($AscString, $addBOM = false)
    Local $BufferSize = StringLen($AscString) * 2
    Local $FullUniStr = DllStructCreate("byte[" & $BufferSize + 2 & "]")
    Local $Buffer = DllStructCreate("byte[" & $BufferSize & "]", DllStructGetPtr($FullUniStr) + 2)
    Local $Return = DllCall("Kernel32.dll", "int", "MultiByteToWideChar", _
        "int", 0, _
        "int", 0, _
        "str", $AscString, _
        "int", StringLen($AscString), _
        "ptr", DllStructGetPtr($Buffer, 1), _
        "int", $BufferSize)
    DllStructSetData($FullUniStr, 1, 0xFF, 1)
    DllStructSetData($FullUniStr, 1, 0xFE, 2)
    If $addBOM then
        Return DllStructGetData($FullUniStr, 1)
    Else
        Return DllStructGetData($Buffer, 1)
    Endif
 EndFunc


 Func _UnicodeToANSI($sString)
	Local Const $SF_ANSI = 1, $SF_UTF8 = 4
	Return BinaryToString(StringToBinary($sString, $SF_ANSI), $SF_UTF8)
EndFunc

ConsoleWrite(($state))
$cv = StringToBinary("TP Há»“ ChÃ­ Minh")
If(StringToBinary( $state )== $cv) Then
   ConsoleWrite($state)
EndIf