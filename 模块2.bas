Attribute VB_Name = "模块2"
Option Explicit

Function makeData(tableItems() As TableDefinition)
    
    Dim ws As Worksheet
    ' 设置目标工作表
    Set ws = ThisWorkbook.Sheets("data")
    
    'データ行番号
    g_dataRowNo = 11

    
End Function



Function init()
    'NUMBER(1)現在の値
    g_number1Val = 1
    'NUMBER(2)現在の値
    g_number2Val = 10
    
    'CHAR(1)現在の値
    g_char1Val = "A"
    'CAHR(2)現在の値
    g_char2Val = "A1"
    
    'VARCHAR(1)現在の値
    g_varChar1Val = "A"
    'VARCAHR(2)現在の値
    g_varChar2Val = "A1"
    
    'Column番号
    g_dataColumn = 1
    
    'Date
    g_dateVal = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    
    
End Function

    
Function editExcel(tableItems() As TableDefinition)
    Dim i As Integer
    Dim ws As Worksheet
    ' 设置目标工作表
    Set ws = ThisWorkbook.Sheets("data")
    ' 按照表变量属性生成数据
    For i = 1 To UBound(tableItems)
        ws.Cells(g_dataRowNo, i + 3).value = tableItems(i).CreateDataValue
    Next i
End Function



