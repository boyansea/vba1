Attribute VB_Name = "ģ��1"
' �����ṹ����
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

'�ǩ`���з���
Public g_dataRowNo As Integer
'DataKey
Public g_dataKey As String
'ȫ�ǥǩ`������
Public g_charFullNo As Integer

'DataColumn
Public g_dataColumn As String

'NUMBER(1)�F�ڤ΂�
Public g_number1Val As Integer
'NUMBER(2)�F�ڤ΂�
Public g_number2Val As Integer

'CHAR(1)�F�ڤ΂�
Public g_char1Val As String
'CAHR(2)�F�ڤ΂�
Public g_char2Val As String

'VARCHAR(1)�F�ڤ΂�
Public g_varChar1Val As String
'VARCAHR(2)�F�ڤ΂�
Public g_varChar2Val As String

'DATE�F�ڤ΂�
Public g_dateVal As Date

'NUMBERС���Х�����
Public g_numberDecByte As Integer
