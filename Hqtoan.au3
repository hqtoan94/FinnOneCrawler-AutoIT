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

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

#include <Excel.au3>
#include <File.au3>

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Auto Crawl Data", 376, 200, 192, 124)

$Label1 = GUICtrlCreateLabel("Signed From:", 24, 80, 98, 24)
GUICtrlSetFont($Label1, 13, 400, 0, "Cambria")

$Label2 = GUICtrlCreateLabel("Signed To:", 24, 112, 78, 24)
GUICtrlSetFont($Label2, 13, 400, 0, "Cambria")

$signedfrom = GUICtrlCreateInput("", 144, 80, 137, 21)
$signedto = GUICtrlCreateInput("", 144, 112, 137, 21)

$Label3 = GUICtrlCreateLabel("Account 2:", 24, 16, 81, 24)
GUICtrlSetFont($Label3, 13, 400, 0, "Cambria")

$Label6 = GUICtrlCreateLabel("0", 320, 64, 78, 24)
GUICtrlSetFont($Label6, 13, 400, 0, "Cambria")

$username = GUICtrlCreateInput("KI000122", 144, 16, 137, 21)
$password = GUICtrlCreateInput("Mongdiep123!", 144, 40, 137, 21)

$StartButton = GUICtrlCreateButton("Start", 300, 100, 65, 41)
$ExitButton = GUICtrlCreateButton("Exit", 300, 150, 65, 41)

$combo = GUICtrlCreateCombo("", 144, 150, 137, 21, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData($combo, "Approval Confirmation|Dedupe Referral|Detail Data Entry|Detail Data Entry Quality Check|Detail Policy Referral|Disbursal Detail|Disbursal Initiation|Document Collection|Document Compliance|FI Bank|FI Completion|FI Initiation|FI Verification Detail|Financial Analysis|Phone Verification|Policy Referral|Post Sanc Doc|Pre Disbursal Document Verification|Quick Data Entry|Quick Data Entry Quality Check|Re-consideration|Reject Review|Scoring Referral|Stage Reversal|Underwriting|WAITING STAGE PRE APPROVAL CREDIT")
#EndRegion ### END Koda GUI section ###

Local $arrayInf[1][10] = [["Full Name", "Gender", "Age", "ID Card Number", "Phone", "State", "Stage", "Scheme", "Company", "Income"]]
Local $row = 1

Local $process = ProcessList("EXCEL.EXE")
For $i = 1 To $process[0][0]
   ProcessClose($process[$i][1])
Next

Local $oExcel = _Excel_Open()
Local $oWorkbook = _Excel_BookNew($oExcel)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error creating new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; Write something into cell A1

_Excel_BookSaveAs($oWorkbook, @ScriptDir & "\_Excel.xlsx", Default, True)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error saving workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $arrayInf, "A" & $row)
$row = $row + 1

_Excel_BookSave($oWorkbook)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error saving workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

GUISetState(@SW_SHOW)
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		 Case $GUI_EVENT_CLOSE
			_Excel_BookSave($oWorkbook)
			_Excel_BookClose($oWorkbook)
			Exit 0
		 Case $StartButton
			If(Firefox(GUICtrlRead($signedfrom), GUICtrlRead($signedto), GUICtrlRead($username), GUICtrlRead($password)) = 2) Then
			   Sleep(1000)
			   Firefox(GUICtrlRead($signedfrom), GUICtrlRead($signedto), GUICtrlRead($username), GUICtrlRead($password))
			EndIf
			MsgBox(64, "Done", "Done")
			Exit
		 Case $ExitButton
			_Excel_BookSave($oWorkbook)
			_Excel_BookClose($oWorkbook)
			Exit 0
	EndSwitch
 WEnd

Func Firefox($signed, $signedto, $username, $password)

Local Const $locationArray = ["TP Há»“ ChÃ­ Minh", "Tá»‰nh BÃ¬nh DÆ°Æ¡ng", "Tá»‰nh Long An", "Tá»‰nh Äá»“ng Nai"]

If _FFStart("https://cps.fecredit.com.vn/finnsso/gateway/SSOGateway?requestID=7000003", Default, 2) Then
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
		 _FFCmd(".getElementsByName('frameForwardToApp')[0].contentWindow.document.getElementsByName('contents')[0].contentWindow.document.getElementsByClassName('menuRow')[2].click()")
		 Sleep(500)
		 _FFCmd(".getElementsByName('frameForwardToApp')[0].contentWindow.document.getElementsByName('contents')[0].contentWindow.document.getElementsByClassName('menu1')[1].getElementsByTagName('div')[0].click()")
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
	  If($signed <> "") Then
		 _FFSetValue($signed, "signed", "name")
	  EndIf
	  If($signedto <> "") Then
		 _FFSetValue($signedto, "signedTo", "name")
	  EndIf

	  $comboRead = GUICtrlRead($combo)
	  If($comboRead <> "") Then
		 Switch $comboRead
		 Case "Approval Confirmation"
			_FFSetValue("APP_CONF", "selActivityId", "name")
		 Case "Dedupe Referral"
		 _FFSetValue("DUR", "selActivityId", "name")
		 Case "Detail Data Entry"
		 _FFSetValue("BDE", "selActivityId", "name")
		 Case "Detail Data Entry Quality Check"
		 _FFSetValue("DDEQC", "selActivityId", "name")
		 Case "Detail Policy Referral"
		 _FFSetValue("DPOR", "selActivityId", "name")
		 Case "Disbursal Detail"
		 _FFSetValue("DISBDTL", "selActivityId", "name")
		 Case "Disbursal Initiation"
		 _FFSetValue("DII", "selActivityId", "name")
		 Case "Document Collection"
		 _FFSetValue("DOC", "selActivityId", "name")
		 Case "Document Compliance"
		 _FFSetValue("DOC_COM", "selActivityId", "name")
		 Case "FI Bank"
		 _FFSetValue("FIB", "selActivityId", "name")
		 Case "FI Completion"
		 _FFSetValue("FIC", "selActivityId", "name")
		 Case "FI Initiation"
		 _FFSetValue("FII", "selActivityId", "name")
		 Case "FI Verification Detail"
		 _FFSetValue("FIV", "selActivityId", "name")
		 Case "Financial Analysis"
		 _FFSetValue("FA", "selActivityId", "name")
		 Case "Phone Verification"
		 _FFSetValue("PHV", "selActivityId", "name")
		 Case "Policy Referral"
		 _FFSetValue("POR", "selActivityId", "name")
		 Case "Post Sanc Doc"
		 _FFSetValue("PDOC", "selActivityId", "name")
		 Case "Pre Disbursal Document Verification"
		 _FFSetValue("DOV", "selActivityId", "name")
		 Case "Quick Data Entry"
		 _FFSetValue("QDE", "selActivityId", "name")
		 Case "Quick Data Entry Quality Check"
		 _FFSetValue("QDEQC", "selActivityId", "name")
		 Case "Re-consideration"
		 _FFSetValue("RECON", "selActivityId", "name")
		 Case "Reject Review"
		 _FFSetValue("REJ", "selActivityId", "name")
		 Case "Scoring Referral"
		 _FFSetValue("SRR", "selActivityId", "name")
		 Case "Stage Reversal"
		 _FFSetValue("STG", "selActivityId", "name")
		 Case "Underwriting"
		 _FFSetValue("UND", "selActivityId", "name")
		 Case "WAITING STAGE PRE APPROVAL CREDIT"
		 _FFSetValue("WSPCA", "selActivityId", "name")
		 EndSwitch
	  EndIf

	  _FFCmd(".getElementsByName('btnSearch')[0].click()")

	  #cs
		 user form
	  #ce
	  If(_FFLoadWait()) Then
		 $count = _FFCmd(".getElementById('selPageIndex').getElementsByTagName('OPTION').length")
		 $i = $count - 1
		 PagePrev($i)

		 Do
			#cs
			   start crawl
			#ce
			For $j = 4 To 23
			   _Excel_BookSave($oWorkbook)
			   _FFTabSetSelected("Enquiry Screen", "label")
			   Sleep(1000)
			   _FFCmd(".getElementsByTagName('a')[" & $j & "].click()")
			   If(_FFLoadWait()) Then
				  #cs
					 Click QDE
				  #ce
				  _FFTabSetSelected(1, "index")
				  Sleep(2000)

				  $stage = Refactor(_FFCmd(".getElementsByName('trackingInterfaceDetailAF')[0].getElementsByClassName('BORDERATTRIBUTES')[0].getElementsByTagName('TR')[4].getElementsByTagName('TD')[1].innerText"))
				  $scheme = Refactor(_FFCmd(".getElementsByName('trackingInterfaceDetailAF')[0].getElementsByClassName('BORDERATTRIBUTES')[0].getElementsByTagName('TR')[8].getElementsByTagName('TD')[1].innerText"))

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
								 $state = Refactor(_FFCmd(".getElementById('selState').value"))

								 If(StringCompare($state, $locationArray[0], $STR_NOCASESENSEBASIC) = 0 Or StringCompare($state, $locationArray[1], $STR_NOCASESENSEBASIC) = 0 Or StringCompare($state, $locationArray[2], $STR_NOCASESENSEBASIC) = 0 Or StringCompare($state, $locationArray[3], $STR_NOCASESENSEBASIC) = 0) Then
									$username = _FFCmd(".getElementById('txtLName').value") & " " & _FFCmd(".getElementById('txtMName').value") & " " & _FFCmd(".getElementById('txtFName').value")

									$gender = _FFCmd(".getElementById('selSex').value")
									$age = _FFCmd(".getElementById('txtAge').value")
									$cardid = "'" & _FFCmd(".getElementById('txtTINNo').value")
									$phone = _FFCmd(".getElementById('txtMobile').value")

									_FFCmd(".getElementById('apy_b1i2').click()")
									If(_FFLoadWait()) Then

									   $company = Refactor(_FFCmd(".getElementById('txtCompnayName').value"))
									   If(StringCompare($company, "") = 0) Then
										  $company = Refactor(_FFCmd(".getElementById('txtOtherEmpName').value"))
									   EndIf

									   _FFCmd(".getElementById('apy_b1i4').click()")

									   If(_FFLoadWait()) Then
										  $countType = _FFCmd(".getElementsByClassName('BORDERATTRIBUTES')[1].getElementsByClassName('LISTTABLEDATA').length")
										  $income = ""
										  For $k = 0 To $countType - 1
											 Local $type = _FFCmd(".getElementsByClassName('BORDERATTRIBUTES')[1].getElementsByClassName('LISTTABLEDATA')[" & $k & "].getElementsByTagName('TD')[1].innerText")
											 If $type = "Thu nháº­p tá»« lÆ°Æ¡ng cá»§a KhÃ¡ch hÃ ngÂ " Then
												$income = Refactor(_FFCmd(".getElementsByClassName('BORDERATTRIBUTES')[1].getElementsByClassName('LISTTABLEDATA')[" & $k & "].getElementsByTagName('TD')[2].innerText"))
												ExitLoop
											 EndIf
										  Next
									   EndIf
									EndIf

									Local $arrayInf2[1][10] = [[_UnicodeToANSI($username), _UnicodeToANSI($gender), _UnicodeToANSI($age), _UnicodeToANSI($cardid), _UnicodeToANSI($phone), _UnicodeToANSI($state), _UnicodeToANSI($stage), _UnicodeToANSI($scheme), _UnicodeToANSI($company), _UnicodeToANSI($income)]]

									_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $arrayInf2, "A" & $row)
									$row = $row + 1
									GUICtrlSetData($Label6, $row - 1)
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
			$i = $i - 1
			If($i = 0) Then
			   ExitLoop
			Else
			   ConsoleWrite($i)
			   PagePrev($i)
			EndIf
		 Until $i = $count

	  EndIf
   EndIf
   Sleep(2000)
   _Excel_BookSave($oWorkbook)
	_FFDisConnect()
 EndIf

EndFunc

Func PagePrev($pagenum)
   _FFSetValue($pagenum, "selPageIndex", "name")
   _FFCmd(".getElementById('selPageIndex').onchange()")

   If(_FFLoadWait()) Then
   EndIf
   Sleep(1000)
EndFunc

 Func Refactor($str)
   If(StringCompare($str, "_FFCmd_Err") = 0) Then
	  Return ""
   EndIf
   Return $str
 EndFunc

Func _UnicodeToANSI($sString)
	Local Const $SF_ANSI = 1, $SF_UTF8 = 4
	Return BinaryToString(StringToBinary($sString, $SF_ANSI), $SF_UTF8)
 EndFunc