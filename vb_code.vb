Sub CSVFileimport()
Dim file_path As Variant
Dim file_name As Variant
  Dim yr As String
    Dim mnt As String
    Dim DATT As String
    DATT = "EXPRESS"
      yr = Year(Date)
       mnt = Month(Date)
          file_name = mnt + yr
         '' file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS"
''file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS\092018.csv"
Workbooks.Open Filename:="D:\docs\logs\MIDAS\SEP\EXPRESS\" & file_name & ".csv", Origin:=xlWindows
    Range("A1").Select
    rowBname = Selection.End(xlDown).Row
     Range("A1:AA" & rowBname).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = (DATT)
    ActiveSheet.Paste
	end sub
	
	--------------------------------------updated as on 04:01-----------------------------
	
	
	
	
	Sub CSVFileimport()
Dim file_path As Variant
Dim file_name As Variant
  Dim yr As String
    Dim mnt As String
    Dim DATT As String
    DATT = "EXPRESS"
      yr = Year(Date)
       mnt = Month(Date)
          file_name = mnt + yr
         '' file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS"
''file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS\092018.csv"
Workbooks.Open Filename:="D:\docs\logs\MIDAS\SEP\EXPRESS\" & file_name & ".csv", Origin:=xlWindows
    Range("A1").Select
    rowBname = Selection.End(xlDown).Row
     Range("A1:AA" & rowBname).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = (DATT)
    ActiveSheet.Paste
    ''Columns ("A2:S2")
    
   ''logic to delete the All the  columns leaving D H N O R
   
    Columns("A:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
   Selection.Delete Shift:=xlToLeft
   Columns("F:F").Select
   Selection.Delete Shift:=xlToLeft
    
	
	
	
	
	
	
	
	
	------------------------------------------
	
	
	
	
Sub CSVFileimport()
Dim file_path As Variant
Dim file_name As Variant

Dim copy_from_wkb As Workbook
Dim copy_to_wkb As Workbook
Dim copy_from_wks As Worksheet
Dim copy_to_wks As Worksheet
Dim header_file As Variant
ChDir "D:\docs\logs\MIDAS\SEP\"
  
  Dim yr As String
    Dim mnt As String
    Dim DATT As String
    DATT = "EXPRESS"
      yr = Year(Date)
       mnt = Month(Date)
          file_name = mnt + yr
         '' file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS"
''file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS\092018.csv"
Workbooks.Open Filename:="D:\docs\logs\MIDAS\SEP\EXPRESS\" & file_name & ".csv", Origin:=xlWindows
    Range("A1").Select
    rowBname = Selection.End(xlDown).Row
     Range("A1:AA" & rowBname).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = (DATT)
    ActiveSheet.Paste
    ''Columns ("A2:S2")
    
    header_file = "mmyyyy_header"
    Set copy_to_wkb = ThisWorkbook
    Set copy_from_wkb = Workbooks.Open("mmyyyy_header.xls")
    Set copy_to_wks = copy_to_wkb.Sheets("EXPRESS")
    Set copy_from_wks = copy_from_wkb.Sheets
    
    
    copy_from_wks.Columns("A:S").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
   '' copy_to_wks.Copy(
    ''copy_to_wks.Columns("A:s").Paste Shift:=xlToDown
    
    
    
    
    
   ''logic to delete the All the  columns leaving D H N O R
   
   '' Columns("A:C").Select
    ''Selection.Delete Shift:=xlToLeft
    ''Columns("B:D").Select
   '' Selection.Delete Shift:=xlToLeft
   '' Columns("C:G").Select
   '' Selection.Delete Shift:=xlToLeft
   '' Columns("E:F").Select
   ''Selection.Delete Shift:=xlToLeft
   ''Columns("F:F").Select
  '' Selection.Delete Shift:=xlToLeft
    
    
''    With ActiveSheet.QueryTables.Add(Connection:= _
  ''    "TEXT;D:\docs\logs\MIDAS\SEP\EXPRESS\092018.csv", Destination:=Range("$A$1"))
    ''    .Name = "092018"
   ''     .FieldNames = True
   ''     .RowNumbers = False
   ''     .FillAdjacentFormulas = False
   ''     .PreserveFormatting = True
   ''     .RefreshOnFileOpen = False
   ''     .RefreshStyle = xlInsertDeleteCells
   ''     .SavePassword = False
   ''     .SaveData = True
   ''     .AdjustColumnWidth = True
   ''     .RefreshPeriod = 0
   ''     .TextFilePromptOnRefresh = False
   ''     .TextFilePlatform = 437
   ''     .TextFileStartRow = 1
    ''    .TextFileParseType = xlDelimited
  ''      .TextFileTextQualifier = xlTextQualifierDoubleQuote
  ''      .TextFileConsecutiveDelimiter = False
  ''      .TextFileTabDelimiter = False
  ''      .TextFileSemicolonDelimiter = False
   ''     .TextFileCommaDelimiter = True
  ''      .TextFileSpaceDelimiter = False
  ''      .TextFileTrailingMinusNumbers = True
  ''      .Refresh BackgroundQuery:=False
  ''  End With
 ''   ChDir "D:\docs\learning\MIDAS"
 ''   ActiveWorkbook.SaveAs Filename:="D:\docs\learning\MIDAS\dev.xlsm", _
  ''      FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub


-----------------------------------------------------------------------------------------------
updated to cut copy past the header data from one mmyyyy.xls to the active worksheet
											
											
											
											
											
Sub CSVFileimport()
Dim file_path As Variant
Dim file_name As Variant

'Dim copy_from_wkb As Workbook
'Dim copy_to_wkb As Workbook
'Dim copy_from_wks As Worksheet
'Dim copy_to_wks As Worksheet
'Dim header_file As Variant
'ChDir "D:\docs\logs\MIDAS\SEP\"
  
  Dim yr As String
    Dim mnt As String
    Dim DATT As String
    DATT = "EXPRESS"
      yr = Year(Date)
       mnt = Month(Date)
          file_name = mnt + yr
          file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS"
''file_path = "D:\docs\logs\MIDAS\SEP\EXPRESS\092018.csv"
Workbooks.Open Filename:="D:\docs\logs\MIDAS\SEP\EXPRESS\" & file_name & ".csv", Origin:=xlWindows
    Range("A1").Select
    rowBname = Selection.End(xlDown).Row
     Range("A1:AA" & rowBname).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = (DATT)
    ActiveSheet.Paste
    ''Columns ("A2:S2")
    
   header_file = "mmyyyy_header"
   
    Workbooks.Open Filename:="D:\docs\logs\MIDAS\SEP\" & header_file & ".xls"
    Application.CutCopyMode = False
    Selection.Copy
    Windows(file_name).Activate
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("G2").Select
    
    Windows("mmyyyy_header.xls").Activate
    ActiveWindow.Close
    'Workbooks.Close Filename:="D:\docs\logs\MIDAS\SEP\" & header_file & ".xls"
   ' Workbooks.Close Filename:="
 '   Set copy_to_wkb = ThisWorkbook
 '   Set copy_from_wkb = Workbooks.Open("mmyyyy_header.xls")
 '   Set copy_to_wks = copy_to_wkb.ActiveSheet
    '' Set copy_from_wks = copy_from_wkb.Sheets
    
    
    'Range("A1").Select
    'rowAname = Selection.End(xlDown).Row
    'Range("A1:S1" & rowBname).Select
   ' Selection.Copy
    ''copy_to_wks.Rows(A1).Insert Shift:=xlToDown
  '  copy_to_wks.Activate
    ' 'Shift:=xlToDown
    ''ActiveCell = A1
    ''Selection.Insert Shift:=xlToDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
     
    
    
     ''copy_from_wks.Columns("A:S").Select
     ''Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
   '' copy_to_wks.Copy(
    ''copy_to_wks.Columns("A:s").Paste Shift:=xlToDown
    
    
    
    
    
   ''logic to delete the All the  columns leaving D H N O R
   
   ''Columns("A:C").Select
  '' Selection.Delete Shift:=xlToLeft
  '' Columns("B:D").Select
  '' Selection.Delete Shift:=xlToLeft
  '' Columns("C:G").Select
  '' Selection.Delete Shift:=xlToLeft
  '' Columns("E:F").Select
   ''Selection.Delete Shift:=xlToLeft
   ''Columns("F:F").Select
   ''Selection.Delete Shift:=xlToLeft
    
    
''    With ActiveSheet.QueryTables.Add(Connection:= _
  ''    "TEXT;D:\docs\logs\MIDAS\SEP\EXPRESS\092018.csv", Destination:=Range("$A$1"))
    ''    .Name = "092018"
   ''     .FieldNames = True
   ''     .RowNumbers = False
   ''     .FillAdjacentFormulas = False
   ''     .PreserveFormatting = True
   ''     .RefreshOnFileOpen = False
   ''     .RefreshStyle = xlInsertDeleteCells
   ''     .SavePassword = False
   ''     .SaveData = True
   ''     .AdjustColumnWidth = True
   ''     .RefreshPeriod = 0
   ''     .TextFilePromptOnRefresh = False
   ''     .TextFilePlatform = 437
   ''     .TextFileStartRow = 1
    ''    .TextFileParseType = xlDelimited
  ''      .TextFileTextQualifier = xlTextQualifierDoubleQuote
  ''      .TextFileConsecutiveDelimiter = False
  ''      .TextFileTabDelimiter = False
  ''      .TextFileSemicolonDelimiter = False
   ''     .TextFileCommaDelimiter = True
  ''      .TextFileSpaceDelimiter = False
  ''      .TextFileTrailingMinusNumbers = True
  ''      .Refresh BackgroundQuery:=False
  ''  End With
 ''   ChDir "D:\docs\learning\MIDAS"
 ''   ActiveWorkbook.SaveAs Filename:="D:\docs\learning\MIDAS\dev.xlsm", _
  ''      FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub








