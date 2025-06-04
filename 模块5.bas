Attribute VB_Name = "模块5"
'全角满位
Function makeDataType2(tableItems() As TableDefinition)

    Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    
    Dim rangeK As range
    Set rangeK = ws.range("K:K")
    g_dataKey = rangeK.Cells(g_dataRowNo - 16).value
    ' 按照表变量属性生成数据
    For i = 1 To UBound(tableItems)
        If tableItems(i).DataType = "CHAR" Then
            tableItems(i).CreateDataValue = makeCharFullData(tableItems(i))
        End If
        If tableItems(i).DataType = "VARCHAR2" Then
            tableItems(i).CreateDataValue = makeVarChar2Data(tableItems(i))
        End If
        If tableItems(i).DataType = "DATE" Then
            tableItems(i).CreateDataValue = getDateVal()
        End If
        If tableItems(i).DataType = "NUMBER" And tableItems(i).DecimalLength <> 0 Then
            tableItems(i).CreateDataValue = makeNumberNoDecData(tableItems(i))
            tableItems(i).CreateDataValue = tableItems(i).CreateDataValue & "." & Format("1", String(tableItems(i).DecimalLength, "0"))
        End If
        If tableItems(i).DataType = "NUMBER" And tableItems(i).DecimalLength = 0 Then
            tableItems(i).CreateDataValue = makeNumberNoDecData(tableItems(i))
        End If
    Next i
    
    Call editExcel(tableItems)
End Function

'CHAR全角データ作成
Function makeCharFullData(itemInfo As TableDefinition) As String

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    Dim rangeB As range
    Set rangeB = ws.range("B:B")
    Dim rangeK As range
    Set rangeK = ws.range("K:K")
    
    If itemInfo.DataLength = 1 Then
        makeCharFullData = " "
    ElseIf itemInfo.DataLength = 2 Or itemInfo.DataLength = 3 Then
        makeCharFullData = getChar2FullVal(rangeK, rangeB.Cells(3).value)
        rangeB.Cells(3).value = makeCharFullData
        makeCharFullData = makeCharFullData & Space(itemInfo.DataLength - Len(makeCharFullData))
    ElseIf itemInfo.DataLength = 4 Or itemInfo.DataLength = 5 Then
        makeCharFullData = getChar4FullVal(rangeK, rangeB.Cells(3).value)
        rangeB.Cells(3).value = makeCharFullData
        makeCharFullData = makeCharFullData & Space(itemInfo.DataLength - Len(makeCharFullData))
    Else
        Dim fullCharCnt As Integer
        fullCharCnt = Int(itemInfo.DataLength / 2)
        makeCharFullData = g_dataKey & StrConv(Format(g_charFullNo, String(fullCharCnt - 1, "0")), vbWide)
        makeCharFullData = makeCharFullData & Space(itemInfo.DataLength - Len(makeCharFullData))
        g_charFullNo = g_charFullNo + 1
    End If
End Function

Function getChar2FullVal(fromRange As range, val As String) As String
    If val = "" Then
        getChar2FullVal = fromRange.Cells(1).value
        Exit Function
    End If
    Dim foundCell As range
    Set foundCell = fromRange.Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)
    getChar2FullVal = foundCell.Offset(1, 0).value
    If foundCell.Row > 46 Then
        getChar2FullVal = fromRange.Cells(1).value
    End If
End Function

Function getChar4FullVal(fromRange As range, val As String) As String
    If val = "" Then
        getCharNFullVal = fromRange.Cells(1).value
        Exit Function
    End If
    Dim foundCell As range
    Set foundCell = fromRange.Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)
    getCharNFullVal = foundCell.Offset(1, 0).value & foundCell.Offset(1, 0).value
    
End Function

