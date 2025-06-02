Attribute VB_Name = "模块3"
Option Explicit


Function getHalfKeyString(keyNo As Integer) As String
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    getKeyString = Mid(chars, keyNo, 1)
End Function
Function getFullKeyString(keyNo As Integer) As String
    chars = "全角文字按照五十音图"
    getFullKeyString = Mid(chars, keyNo, 1)
End Function


'sheet get cell value
Function setValueToCell(range As range, rowNo As Integer, val As String)
    range.Cells(rowNo).value = val
End Function

'sheet get cell value
Function getValueFromCell(rowNo As Integer) As String
    
End Function

'data重複check
Function isDuplication(rowNo As Integer) As Boolean
    

End Function
