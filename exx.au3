#include <Excel.au3>
#include <MsgBoxConstants.au3>

; Create application object
Local $oExcel = _Excel_Open(False)
If @error Then Exit MsgBox(16, "Excel UDF: _Excel_BookOpen Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Create a new workbook, write some data and save it to @tempdir
; *****************************************************************************
; Create the new workbook
Local $oWorkbook = _Excel_BookNew($oExcel)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error creating new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; Write something into cell A1
_Excel_RangeWrite($oWorkbook, Default, "Test", "A1")
If Not IsObj($oWorkbook) Or ObjName($oWorkbook, 1) <> "_Workbook" Then Exit SetError(1, 0, 0)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error writing to cell 'A1'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; Save the workbook to @tempdir as _Excel.xls (replacing existing file)
_Excel_BookSaveAs($oWorkbook, @ScriptDir & "\_Excel.xlsx", Default, True)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error saving workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; Write something into cell B1
Local $aArray2D[1][5] = [[11, 12, 13, 14, 15]]
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aArray2D, "A" & 1)
_Excel_BookSave($oWorkbook)
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aArray2D, "A" & 2)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error writing to cell 'A1'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; Save changed workbook
_Excel_BookSave($oWorkbook)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error saving workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Workbook has been successfully saved as '" & @TempDir & "\_Excel.xls'.")
