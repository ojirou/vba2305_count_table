Attribute VB_Name = "Module1"
'
'#############################################################################
' �yEXL�z �e�[�u���̃��R�[�h���J�E���g
'
'�@table_count
'#############################################################################
'Option Explicit
Sub CountTable()
Dim SheetName As String
Workbooks.Open ("C:\Users\user\git\excel_vba\data_table\list_sample.xlsx")
SheetName = "Sheet1"
Dim rowCount As Long
    rowCount = Worksheets(SheetName).ListObjects("Table1").ListRows.Count
    MsgBox ("�e�[�u���̃��R�[�h���́w" & rowCount & "�x�ł��B")
End Sub
