Attribute VB_Name = "ģ��2"
Option Explicit

Function makeData(tableItems() As TableDefinition)
    
    Dim ws As Worksheet
    ' ����Ŀ�깤����
    Set ws = ThisWorkbook.Sheets("data")
    
    '�ǩ`���з���
    g_dataRowNo = 11

    
End Function



Function init()
    'NUMBER(1)�F�ڤ΂�
    g_number1Val = 1
    'NUMBER(2)�F�ڤ΂�
    g_number2Val = 10
    
    'CHAR(1)�F�ڤ΂�
    g_char1Val = "A"
    'CAHR(2)�F�ڤ΂�
    g_char2Val = "A1"
    
    'VARCHAR(1)�F�ڤ΂�
    g_varChar1Val = "A"
    'VARCAHR(2)�F�ڤ΂�
    g_varChar2Val = "A1"
    
    'Column����
    g_dataColumn = 1
    
    'Date
    g_dateVal = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    
    
End Function

    
Function editExcel(tableItems() As TableDefinition)
    Dim i As Integer
    Dim ws As Worksheet
    ' ����Ŀ�깤����
    Set ws = ThisWorkbook.Sheets("data")
    ' ���ձ����������������
    For i = 1 To UBound(tableItems)
        ws.Cells(g_dataRowNo, i + 3).value = tableItems(i).CreateDataValue
    Next i
End Function



