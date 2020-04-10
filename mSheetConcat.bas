Attribute VB_Name = "mSheetConcat"
Option Explicit

Private Function SheetExistence( _
  ByRef wbkActive As Workbook, _
  ByVal strSheetNameToFind As String, _
  ByVal blnSheetExists As Boolean) As Boolean
  
  

  Dim objSheet As Object
    For Each objSheet In wbkActive.Worksheets
      If strSheetNameToFind = objSheet.Name _
      And blnSheetExists = False Then
        SheetExistence = True
        Exit Function
      End If
    Next objSheet

End Function

Private Function addset_sht( _
  ByRef wbkActive As Workbook, _
  ByVal strSheetName As String) As Worksheet

  Dim blnSheetExists As Boolean

  blnSheetExists = SheetExistence( _
    wbkActive, _
    strSheetName, _
    False)

  With wbkActive
    Select Case blnSheetExists
      Case True
        Set addset_sht = .Sheets(strSheetName)
        Exit Function
    End Select
    ' If the sub goes here, then sheet is note exists
    'and we will create it
    Dim shtNew As Worksheet
    Set shtNew = .Worksheets.Add(After:=.Worksheets(.Worksheets.Count))
    shtNew.Name = strSheetName
    Set addset_sht = shtNew
  End With

End Function

Public Sub ActiveSheetsConcat()
  
  On Error GoTo ErrorHandler
  Application.Calculation = xlManual
  
  Dim dangerText As String
  Dim clsModeToast As cRuleToast
  Set clsModeToast = New cRuleToast

  Dim shtTotal As Worksheet
  Dim rngDataPaste As Range
  Dim rngDataCopy As Variant
  Dim rngShtName As Range
  Dim strNameShtWork As String
  Dim lastrow As Long
  Dim lastrow2 As Long

  Const NAMESHTTOTAL As String = "Total"
  Const DATACOPYCOLUMNS As Integer = 30

  Dim wbkActive As Workbook
  Dim wbkNewBook As Workbook
  Dim blnIsNewBookCreated As Boolean
  Set wbkActive = ActiveWorkbook
  ' We need to do check of an extension
  '
 
  With wbkActive
    Select Case True
      Case .FileFormat = xlOpenXMLWorkbookMacroEnabled _
      Or .FileFormat = xlWorkbookDefault _
      Or .FileFormat = xlExcel12
        Set wbkNewBook = wbkActive
        blnIsNewBookCreated = False

      Case .FileFormat = xlWorkbookNormal _
      Or .FileFormat = xlExcel9795 _
      Or .FileFormat = xlExcel8 _
      Or .FileFormat = xlExcel7 _
      Or .FileFormat = xlExcel5 _
      Or .FileFormat = xlExcel4 _
      Or .FileFormat = xlExcel3 _
      Or .FileFormat = xlExcel2FarEast _
      Or .FileFormat = xlExcel2
        'Important note: we will recreate whole workbook
        'to handle limit in 65536 rows in older versions of excel.
        Set wbkNewBook = Workbooks.Add
        blnIsNewBookCreated = True
        wbkActive.Activate
    End Select
  End With

  Set shtTotal = addset_sht(wbkNewBook, NAMESHTTOTAL)


  Dim shtWork As Worksheet
  Dim rngFirstCell As Range
  For Each shtWork In wbkActive.Worksheets

    With shtWork
      If NAMESHTTOTAL = .Name Then GoTo nexti

      Set rngFirstCell = .Range("A1")
      lastrow = rngFirstCell.CurrentRegion.Rows.Count
      strNameShtWork = .Name
      rngDataCopy = .Range(rngFirstCell, .Cells(lastrow, DATACOPYCOLUMNS)).Value
    End With

    With shtTotal
      Set rngFirstCell = .Range("A1")
      lastrow2 = rngFirstCell.CurrentRegion.Rows.Count + 1
      Set rngDataPaste = .Range(.Cells(lastrow2, 2), _
        .Cells(lastrow2 + lastrow - 1, DATACOPYCOLUMNS + 1))
      rngDataPaste.Value = rngDataCopy
      Set rngShtName = .Range(.Cells(lastrow2, 1), .Cells(lastrow2 + lastrow - 1, 1))
      rngShtName.Value2 = strNameShtWork
    End With

nexti:
  Next shtWork
  Select Case blnIsNewBookCreated
    Case True
      wbkActive.Close SaveChanges:=False
      wbkNewBook.Activate
  End Select

  clsModeToast.OpenToast enmControlType.ectSuccess

Exit sub
ErrorHandler:
dangerText = Err.Description & " " & Err.Number
clsModeToast.OpenToast enmControlType.ectDanger, dangerText
End Sub
