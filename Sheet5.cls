VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function getTableInfo(lastRow As Long, result() As TableDefinition)
    Dim ws As Worksheet
    Dim i As Long
    Dim ItemNo As Integer
    Dim tableDefs() As TableDefinition
    
    ' 设置目标工作表
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim rangeC As range
    Set rangeC = ws.range("C:C")
    Dim rangeG As range
    Set rangeG = ws.range("G:G")
    Dim rangeH As range
    Set rangeH = ws.range("H:H")
    Dim rangeI As range
    Set rangeI = ws.range("I:I")
    Dim rangeJ As range
    Set rangeJ = ws.range("J:J")
    Dim rangeK As range
    Set rangeK = ws.range("K:K")
    Dim rangeL As range
    Set rangeL = ws.range("L:L")
    
    Dim dataDict As Object
    
    ' 创建 Dictionary
    Set dataDict = CreateObject("Scripting.Dictionary")
    
    ' 遍历 C 列
    ItemNo = 1
    For i = 10 To lastRow
        If Trim(ws.Cells(i, 3).value) <> "" Then
            ' 调整数组大小以匹配有效数据量
            ReDim Preserve tableDefs(1 To ItemNo)
            tableDefs(ItemNo).ItemKey = ItemNo
            tableDefs(ItemNo).ItemId = Trim(rangeC.Cells(i).value)
            tableDefs(ItemNo).DataType = Trim(rangeG.Cells(i).value)
            If Trim(rangeH.Cells(i).value) <> "" Then
                tableDefs(ItemNo).DataLength = Trim(rangeH.Cells(i).value)
            Else
                tableDefs(ItemNo).DataLength = 0
            End If
            If Trim(rangeI.Cells(i).value) <> "" Then
                tableDefs(ItemNo).DecimalLength = Trim(rangeI.Cells(i).value)
            Else
                tableDefs(ItemNo).DecimalLength = 0
            End If
            If Trim(rangeJ.Cells(i).value) <> "" Then
                tableDefs(ItemNo).IsPrimaryKey = True
            Else
                tableDefs(ItemNo).IsPrimaryKey = False
            End If
            If Trim(rangeK.Cells(i).value) <> "" Then
                tableDefs(ItemNo).IsNotNull = True
            Else
                tableDefs(ItemNo).IsNotNull = False
            End If
            tableDefs(ItemNo).DefaultValue = Trim(rangeL.Cells(i).value)
            tableDefs(ItemNo).CreateDataOkFlg = False
            tableDefs(ItemNo).ItemNo = ItemNo
            ItemNo = ItemNo + 1
        End If
    Next i
    
    result = tableDefs
    

End Function

' 生成指定长度的随机字符串
Function GenerateRandomString(length As Integer) As String
    Dim i As Integer
    Dim chars As String
    Dim result As String
    
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    result = ""
    
    For i = 1 To length
        result = result & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomString = result
End Function

' 生成指定长度的随机数
Function GenerateRandomNumber(length As Integer) As String
    Dim i As Integer
    Dim result As String
    
    result = ""
    
    For i = 1 To length
        result = result & Int(Rnd() * 10) ' 生成 0-9 的随机数字
    Next i
    
    GenerateRandomNumber = result
End Function

' 生成随机日期时间数据
Function GenerateRandomDate() As String
    GenerateRandomDate = Format(DateAdd("d", Int((365 * Rnd) + 1), Date), "YYYY-MM-DD HH:MM:SS")
End Function

Function makeItemVal(itemInfo As TableDefinition) As String
    Dim randomStr As String: randomStr = ""
    
    ' 确保字符串不重复
    If itemInfo.DataType = "NUMBER" Then
        makeItemVal = GenerateRandomNumber(itemInfo.DataLength)
    End If
    If itemInfo.DataType = "CHAR" Or itemInfo.DataType = "VARCHAR2" Then
        makeItemVal = GenerateRandomString(itemInfo.DataLength)
    End If
    'YYYY-MM-DD HH:MM:SS
    If itemInfo.DataType = "DATE" Then
        makeItemVal = GenerateRandomDate()
    End If
End Function

'主key数据
Function getOnlyPrimaryKeyRecord(itemInfo As TableDefinition) As String
    If itemInfo.IsPrimaryKey Then
        getFullDataRecord = makeItemVal(tableItems(i))
    Else
        getFullDataRecord = ""
    End If
End Function

Sub GenerateAndSaveRandomRecord()
    
    Dim tableItems() As TableDefinition
    

    Dim randomStr As String
    Dim i As Integer
    Dim filePath As String
    Dim fso As Object, txtFile As Object
    
    
    Dim ws As Worksheet
    ' 设置目标工作表
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long
    ' 获取最后一行（确保数据范围）
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' 动态数组大小
    ReDim projects(1 To lastRow - 1)
    
    Call getTableInfo(lastRow, tableItems)
    
    
    '旧データClear
    ThisWorkbook.Sheets("data").range("D10").Resize(1, Columns.Count - 3).ClearContents
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    
    Call init

    'PKNotNullOnly
    g_dataRowNo = 11
    makeDataType0 tableItems
    Call editExcel(tableItems)
        

    '全項目最大桁数
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    Call getTableInfo(lastRow, tableItems)
    g_dataRowNo = 12
    makeDataType1 tableItems
    Call editExcel(tableItems)
    
    '半角文字末尾Space
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    Call getTableInfo(lastRow, tableItems)
    g_dataRowNo = 13
    makeDataType1 tableItems
    
    Dim tempVal As String
    For i = 1 To UBound(tableItems)
        tempVal = tableItems(i).CreateDataValue
        If tableItems(i).DataLength = 1 Or tableItems(i).DataType = "NUMBER" Then
        
        ElseIf tableItems(i).DataType = "CHAR" Or tableItems(i).DataType = "VARCHAR2" Then
            tempVal = Left(tempVal, Len(tempVal) - 1)
            tableItems(i).CreateDataValue = tempVal & " "
        End If
    Next i
    
    Call editExcel(tableItems)
    
    '半角文字中間Space
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    Call getTableInfo(lastRow, tableItems)
    g_dataRowNo = 14
    makeDataType1 tableItems
    
    For i = 1 To UBound(tableItems)
        tempVal = tableItems(i).CreateDataValue
        If tableItems(i).DataLength = 1 Or tableItems(i).DataType = "NUMBER" Then
        
        ElseIf tableItems(i).DataType = "CHAR" Or tableItems(i).DataType = "VARCHAR2" Then
            tableItems(i).CreateDataValue = Left(tempVal, 1) & " " & Mid(tempVal, 3)
        End If
    Next i
    
    Call editExcel(tableItems)
    
    '半角文字先頭Space
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    Call getTableInfo(lastRow, tableItems)
    g_dataRowNo = 15
    makeDataType1 tableItems
    
    For i = 1 To UBound(tableItems)
        tempVal = tableItems(i).CreateDataValue
        If tableItems(i).DataLength = 1 Or tableItems(i).DataType = "NUMBER" Then
        
        ElseIf tableItems(i).DataLength = 2 And (tableItems(i).DataType = "CHAR" Or tableItems(i).DataType = "VARCHAR2") Then
            tableItems(i).CreateDataValue = " " & Mid(tempVal, 2)
        ElseIf tableItems(i).DataType = "CHAR" Or tableItems(i).DataType = "VARCHAR2" Then
            tableItems(i).CreateDataValue = " " & Left(tempVal, 1) & Mid(tempVal, 3, Len(tempVal) - 1)
        End If
    Next i
    
    Call editExcel(tableItems)
    
    
    '半角文字Spaceなし
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    Call getTableInfo(lastRow, tableItems)
    g_dataRowNo = 16
    makeDataType1 tableItems
    Call editExcel(tableItems)
    
    
    '全角文字Spaceなし
    g_charFullNo = 1
    ThisWorkbook.Sheets("numbering").range("B:B,F:F").ClearContents
    Call getTableInfo(lastRow, tableItems)
    g_dataRowNo = 20
    makeDataType2 tableItems
    Call editExcel(tableItems)
    
    ' CSV 文件
    filePath = ThisWorkbook.Path & "\outPut.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(filePath, True)
    
    ' 创建数据字典
    Dim csvLineDict As Object
    Set csvLineDict = CreateObject("Scripting.Dictionary")
    
    ' 主key数据字典
    Dim primaryKeys As Object
    ' 创建 Dictionary 存储唯一主键组合
    Set primaryKeys = CreateObject("Scripting.Dictionary")
    
    ' 数据生成结束flg
    Dim makeDataContinue As Boolean: makeDataContinue = True
    Do
        ' 数据生成
        makeRecord tableItems, primaryKeys, csvLineDict
        ' csv一件数据写出
        writeCsv txtFile, csvLineDict
        csvLineDict.RemoveAll
        
        makeDataContinue = False
        For i = 1 To UBound(tableItems)
            If tableItems(i).CreateDataOkFlg = False Then
                makeDataContinue = True
            End If
        Next i
        
    Loop Until Not makeDataContinue
    
    txtFile.Close
    MsgBox "RandomCsv：" & filePath, vbInformation
End Sub

Function writeCsv(txtFile As Object, csvLineDict As Object)

    Dim csvLine As String: csvLine = ""
    Dim key As Variant
    For Each key In csvLineDict
        csvLine = csvLine & csvLineDict(key) & ","
    Next key
    
    ' 去掉最后一个逗号
    If Len(csvLine) > 0 Then csvLine = Left(csvLine, Len(csvLine) - 1)
        
    ' 打开文件进行写入
    txtFile.WriteLine csvLine
    
End Function

Function makeRecord(tableItems() As TableDefinition, primaryKeys As Object, csvLineDict As Object)

    Dim i As Integer
    Dim randomStr As String
    ' 创建字典对象用于存储唯一值
    Dim allItemDict As Object
    Set allItemDict = CreateObject("Scripting.Dictionary")
    
    ' 合并主key
    Dim primaryKeysStr As String: primaryKeysStr = ""
    
    'ok Data
    Dim isOkData As Boolean: isOkData = False
    ' 按照表变量属性生成数据
    Do
        primaryKeysStr = ""
        csvLineDict.RemoveAll
        allItemDict.RemoveAll
        For i = 1 To UBound(tableItems)
            Do
                isOkData = False
                randomStr = ""
                If tableItems(i).IsPrimaryKey Then
                    '主key项目重新生成
                    randomStr = makeItemVal(tableItems(i))
                    isOkData = isExistsVal(randomStr, allItemDict)
                    primaryKeysStr = primaryKeysStr & randomStr
                ElseIf tableItems(i).IsNotNull Then
                    '非空项目重新生成
                    randomStr = makeItemVal(tableItems(i))
                    isOkData = isExistsVal(randomStr, allItemDict)
                ElseIf tableItems(i).CreateDataOkFlg Then
                    '已经生成过的项目不再生成
                    randomStr = ""
                    isOkData = True
                Else
                    '未生成项目继续生成
                    randomStr = makeItemVal(tableItems(i))
                    isOkData = isExistsVal(randomStr, allItemDict)
                End If
                If tableItems(i).CreateDataOkFlg = False And Not isOkData Then
                    randomStr = ""
                    isOkData = True
                End If
            Loop Until isOkData
            '主key,Not Null项目唯一
            If isOkData And randomStr <> "" Then
                ' 字典追加
                allItemDict.Add randomStr, randomStr
                tableItems(i).CreateDataOkFlg = True
            End If
            
            csvLineDict.Add i, randomStr
        Next i
    Loop Until Not primaryKeys.exists(primaryKeysStr)
    primaryKeys.Add primaryKeysStr, primaryKeysStr

End Function

Function isExistsVal(inVal As String, allItemDict As Object) As Boolean
    If allItemDict.exists(inVal) Then
        isExistsVal = False
        Exit Function
    End If
    isExistsVal = True
End Function


Private Sub CommandButton1_Click()
    Call GenerateAndSaveRandomRecord
End Sub

