Attribute VB_Name = "Module1"
'
'#############################################################################
' 【EXL】 テーブルのレコード数カウント
'
'　table_count
'#############################################################################
'Option Explicit
Sub CountTable()
Dim SheetName As String
Workbooks.Open ("C:\Users\user\git\excel_vba\data_table\list_sample.xlsx")
SheetName = "Sheet1"
Dim rowCount As Long
    rowCount = Worksheets(SheetName).ListObjects("Table1").ListRows.Count
    MsgBox ("テーブルのレコード数は『" & rowCount & "』です。")
End Sub
