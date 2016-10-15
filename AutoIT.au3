#cs
	[CWAutCompFileInfo]
	Company=Personal
	Copyright=Ho Quoc Toan
	Description=Crawl data on finnone
	Version=1.0
	ProductName=Auto Crawl Data
	ProductVersion=1.0.0
#ce

#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#Include <FF.au3>
#Include <Excel.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Auto Crawl Data", 376, 155, 192, 124)
$Label1 = GUICtrlCreateLabel("Signed From:", 24, 80, 98, 24)
GUICtrlSetFont(-1, 13, 400, 0, "Cambria")
$Label2 = GUICtrlCreateLabel("Signed To:", 24, 112, 78, 24)
GUICtrlSetFont(-1, 13, 400, 0, "Cambria")
$signedfrom = GUICtrlCreateInput("15/05/2015", 144, 80, 129, 21)
$signedto = GUICtrlCreateInput("16/05/2015", 144, 112, 129, 21)
$Label3 = GUICtrlCreateLabel("Account 2:", 32, 16, 81, 24)
GUICtrlSetFont(-1, 13, 400, 0, "Cambria")
$username = GUICtrlCreateInput("RTTU0018", 144, 16, 137, 21)
$password = GUICtrlCreateInput("doanTu@9090", 144, 40, 137, 21)
$StartButton = GUICtrlCreateButton("StartButton", 288, 88, 65, 41)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		 Case $GUI_EVENT_CLOSE
			Exit
		 Case $StartButton
			If(Firefox(GUICtrlRead($signedfrom), GUICtrlRead($signedto), GUICtrlRead($username), GUICtrlRead($password)) = 2) Then
			   Sleep(1000)
			   Firefox(GUICtrlRead($signedfrom), GUICtrlRead($signedto), GUICtrlRead($username), GUICtrlRead($password))
			EndIf
			MsgBox(64, "Done", "Done")
			Exit
	EndSwitch
 WEnd

Func Firefox($signed, $signedto, $username, $password)

Local $oExcel = ObjCreate("Excel.Application")
With $oExcel
    .SheetsInNewWorkbook = 1
    .WorkBooks.Add
    .ActiveWorkbook.Worksheets(1).Name = "Foo"
EndWith

Local $array[6] = ["Full name", "Gender", "Age", "ID Cardnumber", "Phone", "State"]
For $i = 0 To 5
   $oExcel.Activesheet.Cells(1,($i+1)).Value = $array[$i]
Next

If _FFStart("https://cps.fecredit.com.vn/finnsso/gateway/SSOGateway", Default, 2) Then
   If(_FFLoadWait()) Then
	  Sleep(2000)
	  Local $FF = _FFWindowGetHandle()
	  Sleep(2000)
	  If(_FFCmd(".location.href") = "https://cpsauthen.fecredit.com.vn/vpn/tmindex.html") Then
		 ControlFocus($FF, "", "")
		 Sleep(2000)
		 _FFSetValue("Lequangbinh", "login", "name")
		 _FFSetValue("LQBvpb89", "passwd", "name")

		 _FFFormSubmit()
		 If(_FFLoadWait()) Then
			Return 2
		 EndIf
	  EndIf
	  ControlFocus($FF, "", "")
	  _FFLinkClick("Click here to login", "text")
	  If(_FFLoadWait()) Then
		 Sleep(1000)
		 _FFSetValue($username, "TxtUID", "name")
		 _FFSetValue($password, "TxtPWD", "name")

		 Sleep(500)
		 _FFCmd(".getElementsByName('DataAction')[0].click()")
	  EndIf
   EndIf

   If(_FFLoadWait()) Then
	  Sleep(2000)
	  If(_FFCmd(".getElementsByName('frmRefresh')[0].getAttribute('action')") = "/finnsso/gateway/SSOGateway?requestID=7000003") Then
		 MsgBox(64, "Warning", "Account " & $username & " has logged in or has changed password")
		 _FFDisConnect()
		 Return 0
	  EndIf

	  Sleep(2000)
	  Send("{ENTER}")
	  Sleep(1000)
	  _FFCmd(".getElementById('btnCAS').click()")

	  If(_FFLoadWait()) Then
		 Sleep(2000)
		 _FFCmd(".getElementsByName('frameForwardToApp')[0].contentWindow.document.getElementsByName('contents')[0].contentWindow.document.getElementsByTagName('div')[0].click()")
		 Sleep(500)
		 _FFCmd(".getElementsByName('frameForwardToApp')[0].contentWindow.document.getElementsByName('contents')[0].contentWindow.document.getElementsByTagName('div')[2].click()")
		 Sleep(500)
	  EndIf
	  #cs
		 Click may cai kia, xong roi den logout
	  #ce
	  If(_FFLoadWait()) Then
		 Sleep(2000)
		 _FFTabSetSelected("FinnOne SSO", "label")
		 Sleep(2000)
		 _FFCmd(".getElementsByName('btnEXIT')[0].click()")
		 Sleep(3000)
		 _FFTabClose("FinnOne SSO", "label")
	  EndIf
	  Sleep(2000)
	  _FFTabSetSelected("Enquiry Screen", "label")
	  Sleep(1000)
	  _FFSetValue("PERSONAL", "selProduct", "name")
	  _FFSetValue($signed, "signed", "name")
	  _FFSetValue($signedto, "signedTo", "name")

	  _FFCmd(".getElementsByName('btnSearch')[0].click()")

	  #cs
		 user form
	  #ce
	  If(_FFLoadWait()) Then
		 $count = _FFCmd(".getElementsByTagName('option').length") - 38
		 $i = 0

		 Do
			#cs
			   start crawl
			#ce
			For $j = 4 To 23
			   _FFTabSetSelected("Enquiry Screen", "label")
			   Sleep(1000)
			   _FFCmd(".getElementsByTagName('a')[" & $j & "].click()")
			   If(_FFLoadWait()) Then
				  #cs
					 Click QDE
				  #ce
				  _FFTabSetSelected(1, "index")
				  Sleep(2000)
				  If(_FFCmd(".getElementsByTagName('a')[0].getAttribute('href').startsWith('Activity.los?activity=QDE')") = 1) Then
					 _FFCmd(".getElementsByTagName('a')[0].click()")

					 If(_FFLoadWait()) Then
						Sleep(1000)
						_FFCmd(".getElementById('apy_b0i2').click()")
						If(_FFLoadWait()) Then
						   Sleep(3000)
						   If(_FFCmd(".getElementsByTagName('a')[1].getAttribute('href').startsWith('QDEPersonal.los?')") = 1) Then
							  _FFCmd(".getElementsByTagName('a')[1].click()")
							  If(_FFLoadWait()) Then
								 Sleep(3000)
								 $state = _FFCmd(".getElementById('selState').value")
								 If($state == "TP Há»“ ChÃ­ Minh") Then
									$username = _FFCmd(".getElementById('txtLName').value") & " " & _FFCmd(".getElementById('txtMName').value") & " " & _FFCmd(".getElementById('txtFName').value")

									$gender = _FFCmd(".getElementById('selSex').value")
									$age = _FFCmd(".getElementById('txtAge').value")
									$cardid = _FFCmd(".getElementById('txtIdNum').value")
									If($cardid = "") Then
									   $cardid = _FFCmd(".getElementById('txtTINNo').value")
									EndIf
									$phone = _FFCmd(".getElementById('txtMobile').value")

									$oExcel.Activesheet.Cells(1, 1).Value = $username
									$oExcel.Activesheet.Cells(1, 2).Value = $gender
									$oExcel.Activesheet.Cells(1, 3).Value = $age
									$oExcel.Activesheet.Cells(1, 4).Value = $cardid
									$oExcel.Activesheet.Cells(1, 5).Value = $phone
									$oExcel.Activesheet.Cells(1, 5).Value = $state
									;FileWriteLine($hFileOpen, _UnicodeToANSI($username & "," & $gender & "," & $age & "," & $cardid & "," & $phone & "," & $state))
								 EndIf

								 Sleep(1000)
								 _FFCmd(".getElementsByName('Image3')[0].click()")
								 Sleep(1000)
								 _FFTabSetSelected("Enquiry Screen", "label")
							  EndIf
						   EndIf
						EndIf
					 EndIf
				  EndIf
			   EndIf
			Next
			#cs
			   end crawl
			#ce
			$i = $i + 1
			If($i = $count) Then
			   ExitLoop
			Else
			   If(_FFCmd(".getElementsByTagName('a')[26].getAttribute('href').startsWith('javascript:nextPage1();')") = 1) Then
				  _FFCmd(".getElementsByTagName('a')[26].click()")
				  If(_FFLoadWait()) Then
				  Endif
				  Sleep(1000)
			   EndIf
			EndIf
		 Until $i = $count

	  EndIf
   EndIf
   Sleep(2000)
   _FFDisConnect()
   $oExcel.Visible = 1
 EndIf
 EndFunc

 Func _UnicodeToANSI($sString)
	Local Const $SF_ANSI = 1, $SF_UTF8 = 4
	Return BinaryToString(StringToBinary($sString, $SF_ANSI), $SF_UTF8)
EndFunc