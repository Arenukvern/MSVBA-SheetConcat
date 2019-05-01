Attribute VB_Name = "MainModule"
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
    Set shtNew = .Worksheets.Add
    shtNew.Name = strSheetName
    Set addset_sht = shtNew
  End With
    
End Function

Public Sub ActiveSheetsConcat()
  Application.Calculation = xlManual
  
  Dim shtTotal As Worksheet
  Dim rngDataPaste As Range
  Dim rngDataCopy As Variant
  Dim rngShtName As Range
  Dim strNameShtWork As String
  Dim lastrow As Long
  Dim lastrow2 As Long
  
  Const NAMESHTTOTAL As String = "Total"
  Const DATACOPYCOLUMNS As Integer = 30
  'Important note: we will recreate whole workbook
  'to handle limit in 65536 rows in older versions of excel.
  ' There need to do some check of this version, then
  ' we will handle this problem in more accurate way
  Dim wbkActive As Workbook
  Dim wbkNewBook As Workbook
  Set wbkActive = ActiveWorkbook
  Set wbkNewBook = Workbooks.Add
  wbkActive.Activate
  
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
  
  wbkActive.Close SaveChanges:=False
  wbkNewBook.Activate

End Sub
