Attribute VB_Name = "模块4"
Option Explicit
'PKNotNullOnly
Function makeDataType0(tableItems() As TableDefinition)

    Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    
    Dim rangeI As range
    Set rangeI = ws.range("I:I")
    g_dataKey = rangeI.Cells(g_dataRowNo - 10).value
    ' 按照表变量属性生成数据
    For i = 1 To UBound(tableItems)
        If tableItems(i).DataType = "CHAR" Then
            tableItems(i).CreateDataValue = makeCharData(tableItems(i))
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
    For i = 1 To UBound(tableItems)
        If Not tableItems(i).IsPrimaryKey And Not tableItems(i).IsNotNull Then
            tableItems(i).CreateDataValue = Space(tableItems(i).DataLength)
        End If
    Next i
    
End Function

'半角满位
Function makeDataType1(tableItems() As TableDefinition)

    Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    
    Dim rangeI As range
    Set rangeI = ws.range("I:I")
    g_dataKey = rangeI.Cells(g_dataRowNo - 10).value
    ' 按照表变量属性生成数据
    For i = 1 To UBound(tableItems)
        If tableItems(i).DataType = "CHAR" Then
            tableItems(i).CreateDataValue = makeCharData(tableItems(i))
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
    
End Function


'CHARデータ作成
Function makeCharData(itemInfo As TableDefinition) As String

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    Dim rangeB As range
    Set rangeB = ws.range("B:B")
    Dim rangeI As range
    Set rangeI = ws.range("I:I")
    
    If itemInfo.DataLength = 1 Then
        makeCharData = getChar1Val(rangeI, rangeB.Cells(2).value)
        rangeB.Cells(2).value = makeCharData
    ElseIf itemInfo.DataLength = 2 Then
        makeCharData = getChar2Val(rangeI, rangeB.Cells(3).value)
        rangeB.Cells(3).value = makeCharData
    ElseIf itemInfo.DataLength = 3 Then
        makeCharData = getChar3Val(rangeI, rangeB.Cells(4).value)
        rangeB.Cells(4).value = makeCharData
    Else
        makeCharData = g_dataKey & Format(itemInfo.ItemNo, String(itemInfo.DataLength - 1, "0"))
    End If
End Function

'VARCHAR2データ作成
Function makeVarChar2Data(itemInfo As TableDefinition) As String

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    Dim rangeD As range
    Set rangeD = ws.range("D:D")
    Dim rangeI As range
    Set rangeI = ws.range("I:I")
    
    If itemInfo.DataLength = 1 Then
        makeVarChar2Data = getChar1Val(rangeI, rangeD.Cells(2).value)
        rangeD.Cells(2).value = makeVarChar2Data
    ElseIf itemInfo.DataLength = 2 Then
        makeVarChar2Data = getChar2Val(rangeI, rangeD.Cells(3).value)
        rangeD.Cells(3).value = makeVarChar2Data
    ElseIf itemInfo.DataLength = 3 Then
        makeVarChar2Data = getChar3Val(rangeI, rangeD.Cells(4).value)
        rangeD.Cells(4).value = makeVarChar2Data
    Else
        makeVarChar2Data = g_dataKey & Format(itemInfo.ItemNo, String(itemInfo.DataLength - 1, "0"))
    End If
End Function

Function getChar1Val(fromRange As range, val As String) As String
    If val = "" Then
        getChar1Val = fromRange.Cells(1).value
        Exit Function
    End If
    Dim foundCell As range
    Set foundCell = fromRange.Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)
    getChar1Val = foundCell.Offset(1, 0).value
End Function

Function getChar2Val(fromRange As range, val As String) As String
    If val = "" Then
        getChar2Val = g_dataKey & fromRange.Cells(1).value
        Exit Function
    End If
    Dim foundCell As range
    Set foundCell = fromRange.Find(What:=Right(val, 1), LookIn:=xlValues, LookAt:=xlWhole)
    getChar2Val = Left(val, 1) & foundCell.Offset(1, 0).value
End Function
Function getChar3Val(fromRange As range, val As String) As String
    If val = "" Then
        getChar3Val = g_dataKey & String(fromRange.Cells(1).value, 2)
        Exit Function
    End If
    Dim foundCell As range
    Set foundCell = fromRange.Find(What:=Right(val, 1), LookIn:=xlValues, LookAt:=xlWhole)
    getChar3Val = Left(val, 1) & String(foundCell.Offset(1, 0).value, 2)
End Function

'Date
Function getDateVal() As String
    getDateVal = DateAdd("d", 1, DateAdd("n", 1, g_dateVal))
    g_dateVal = getDateVal
End Function

'NUMBER小数なしデータ作成
Function makeNumberNoDecData(itemInfo As TableDefinition) As String

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("numbering")
    Dim rangeF As range
    Set rangeF = ws.range("F:F")
    Dim rangeI As range
    Set rangeI = ws.range("I:I")
    
    If itemInfo.DataLength = 1 Then
        makeNumberNoDecData = getNum1Val(rangeF.Cells(2).value)
        rangeF.Cells(2).value = makeNumberNoDecData
    ElseIf itemInfo.DataLength = 2 Then
        makeNumberNoDecData = getNum2Val(rangeF.Cells(3).value)
        rangeF.Cells(3).value = makeNumberNoDecData
    Else
        makeNumberNoDecData = Format(itemInfo.ItemNo + 10 ^ (itemInfo.DataLength - 1), String(itemInfo.DataLength - 1, "0"))
    End If
End Function
Function getNum1Val(val As String) As String
    If val = "" Then
        getNum1Val = 1
        Exit Function
    End If
    getNum1Val = CStr(CInt(val) + 1)
    If getNum1Val = 10 Then
        getNum1Val = 1
    End If
End Function
Function getNum2Val(val As String) As String
    If val = "" Then
        getNum1Val = 10
        Exit Function
    End If
    getNum1Val = Str(CInt(getNum1Val) + 1)
    If getNum1Val = 100 Then
        getNum1Val = 10
    End If
End Function






