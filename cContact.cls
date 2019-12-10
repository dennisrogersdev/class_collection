VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private contact As New Collection

Private Sub Class_Initialize()
    contact.Add 0, "id"
    contact.Add "", "name"
    contact.Add "", "phone"
End Sub

Public Function getValue( _
    ByVal attrib As String) As Variant
    getValue = contact(attrib)
End Function

Public Function setValue( _
    ByVal attrib As String, _
    ByVal value As Variant)
    contact.Remove attrib
    contact.Add value, attrib
End Function
