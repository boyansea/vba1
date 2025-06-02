Attribute VB_Name = "模块1"
' 定义表结构对象
Type TableDefinition
    ItemNo As Integer
    ItemKey As Integer
    ItemId As String
    DataType As String
    DataLength As Integer
    DecimalLength As Integer
    IsNotNull As Boolean
    IsPrimaryKey As Boolean
    DefaultValue As String
    RandomValue As String
    CreateDataOkFlg As Boolean
    CreateDataValue As String
End Type

'データ行番号
Public g_dataRowNo As Integer
'DataKey
Public g_dataKey As String
'全角データ番号
Public g_charFullNo As Integer

'DataColumn
Public g_dataColumn As String

'NUMBER(1)現在の値
Public g_number1Val As Integer
'NUMBER(2)現在の値
Public g_number2Val As Integer

'CHAR(1)現在の値
Public g_char1Val As String
'CAHR(2)現在の値
Public g_char2Val As String

'VARCHAR(1)現在の値
Public g_varChar1Val As String
'VARCAHR(2)現在の値
Public g_varChar2Val As String

'DATE現在の値
Public g_dateVal As Date

'NUMBER小数バイト数
Public g_numberDecByte As Integer
