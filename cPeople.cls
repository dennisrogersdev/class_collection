VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cls As clsClassDefault

Public Sub setValue( _
    ByVal attrib As String, _
    ByVal value As Variant)
    cls.setValue attrib, value
End Sub

Public Property Get getValue( _
    ByVal attrib As String) As Variant
Attribute getValue.VB_UserMemId = 0
    getValue = cls.getValue(attrib)
End Property

Private Sub Class_Initialize()
    Set cls = New clsClassDefault
    Set cls.Connection = Conn
    cls.Table = "cad_clientes"
    cls.FieldId = "codigo"
    cls.Mapping
End Sub

Public Function getString( _
    Optional ByVal sDelimiter As String = vbTab) As String
    getString = cls.getString(sDelimiter)
End Function

Public Function getById() As ADODB.Recordset
    Set getById = cls.getById
End Function

