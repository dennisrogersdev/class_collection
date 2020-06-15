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
Private values As Collection

Public Function getValue( _
    ByVal attrib As String) As Variant
    getValue = values(attrib)
End Function

Public Function setValue( _
    ByVal attrib As String, _
    ByVal value As Variant)
    values.Remove attrib
    values.Add value, attrib
End Function

Private Sub Class_Initialize()
    Dim cls As New clsClassDefault
    Set cls.Connection = Conn
    cls.Table = "cad_clientes"
    Set values = cls.Mapping
    
    Debug.Print values("nome")
End Sub
