Attribute VB_Name = "ģ��3"
Option Explicit


Function getHalfKeyString(keyNo As Integer) As String
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    getKeyString = Mid(chars, keyNo, 1)
End Function
Function getFullKeyString(keyNo As Integer) As String
    chars = "ȫ�����ְ�����ʮ��ͼ"
    getFullKeyString = Mid(chars, keyNo, 1)
End Function


'sheet get cell value
Function setValueToCell(range As range, rowNo As Integer, val As String)
    range.Cells(rowNo).value = val
End Function

'sheet get cell value
Function getValueFromCell(rowNo As Integer) As String
    
End Function

'data���}check
Function isDuplication(rowNo As Integer) As Boolean
    

End Function
