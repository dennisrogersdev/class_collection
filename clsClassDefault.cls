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

Public Property Get values() As Collection
    Set values = cCollection
End Property

Public Property Let values(ByVal vNewValue As Collection)
    Set cCollection = vNewValue
End Property

Public Sub setValue( _
    ByVal attrib As String, _
    ByVal value As Variant)
    values.Remove attrib
    values.Add value, attrib
End Sub

Public Property Get getValue( _
    ByVal attrib As String) As Variant
    getValue = values(attrib)
End Property

Public Function Mapping()
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim rsField As ADODB.Field

    Set cCollection = New Collection
    
    Set rs = Conn.Execute("SELECT * FROM " & sTable & " LIMIT 0")
    
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

Public Sub buildClass( _
    ByVal strTable As String, _
    ByVal strFieldId As String, _
    ByVal connDB As ADODB.Connection)
    
    sTable = strTable
    sFieldId = strFieldId
    Set Conn = connDB
    
    Call Mapping
End Sub

Public Function getString( _
    Optional ByVal sDelimiter As String = vbTab) As String
    
    Dim sStr As String
    Dim Col As Variant
    
    For Each Col In cCollection
        sStr = sStr & sDelimiter & Col
    Next
    
    getString = Mid(sStr, 2)
End Function

Public Function getById() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    Dim Field As ADODB.Field
    
    Set rs = Conn.Execute("SELECT * FROM " & Table & " WHERE " & sFieldId & "=" & cCollection(sFieldId))
    
    If rs.EOF Then
        Call Mapping
    Else
        For Each Field In rs.Fields
            setValue Field.Name, Field.value
        Next
    End If
    
    Set getById = rs
End Function

Public Function getListAll( _
    Optional ByVal Limit As Long = 0) As ADODB.Recordset
    
    Set getListAll = Conn.Execute("SELECT * FROM " & Table & IIf(Limit > 0, " LIMIT " & Limit, ""))
End Function

Public Function getListFilter( _
    ParamArray Args() As Variant)
    
    Dim arg As Variant
    Dim sFilter As String
    
    For Each arg In Args
        sFilter = sFilter & " AND " & arg
    Next
    
    sFilter = Mid(sFilter, 5)
    
    Set getListFilter = Conn.Execute("SELECT * FROM " & Table & " WHERE " & sFilter)
End Function

Public Function Insert()
    Dim query As String
    Dim rs As ADODB.Recordset
    Dim Field As ADODB.Field
    
    Set rs = Conn.Execute("SELECT * FROM " & Table & " LIMIT 0")
    
    query = ""
    query = query & "INSERT INTO "
    query = query & "   " & sTable & " ("
    
    For Each Field In rs.Fields
        query = query & "   " & Field.Name & ","
    Next
    
    query = Left(query, Len(query) - 1)
    query = query & " ) VALUES ( "
    
    For Each Field In rs.Fields
        'Verificar Field.Type para tratar gravação no banco de dados
        query = query & "   " & values(Field.Name) & ","
    Next
    
    query = Left(query, Len(query) - 1)
    query = query & " )"
    
End Function

