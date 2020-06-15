VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClassDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sTable As String
Private sFieldId As String
Private Conn As ADODB.Connection
Private cCollection As Collection

Public Property Get values() As Collection
    Set values = cCollection
End Property

Public Property Let values(ByVal vNewValue As Collection)
    Set cCollection = vNewValue
End Property

Public Property Get Table() As String
    Table = sTable
End Property

Public Property Let Table(ByVal vNewValue As String)
    sTable = vNewValue
End Property

Public Property Get FieldId() As String
    FieldId = sFieldId
End Property

Public Property Let FieldId(ByVal vNewValue As String)
    sFieldId = vNewValue
End Property

Public Property Set Connection(ByVal vNewValue As ADODB.Connection)
    Set Conn = vNewValue
End Property

Public Function Mapping()
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim rsField As ADODB.Field
    
    
    Set cCollection = New Collection
    
    Set rs = Conn.Execute("SELECT * FROM " & sTable)
    
    For Each rsField In rs.Fields
        Select Case rsField.Type
            Case adBigInt, adNumeric, adDecimal, adDouble, _
                 adSmallInt, adTinyInt, adCurrency, adUnsignedBigInt, _
                 adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, _
                 adInteger
                cCollection.Add 0, rsField.Name
            Case Else
                cCollection.Add "", rsField.Name
        End Select
    Next
    
    Set Mapping = cCollection
End Function

Public Function getById() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    Dim Field As ADODB.Field
    
    Set rs = Conn.Execute("SELECT * FROM " & Table & " WHERE " & sFieldId & "=" & cCollection(sFieldId))
    
    If rs.EOF Then
        Call Mapping
    Else
        For Each Field In rs.Fields
            cCollection.Remove Field.Name
            cCollection.Add Field.value, Field.Name
        Next
    End If
    
    Set getById = rs
End Function
