VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "test"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   315
      Left            =   675
      TabIndex        =   4
      Top             =   1425
      Width           =   1140
   End
   Begin VB.TextBox txtPhone 
      Height          =   360
      Left            =   375
      TabIndex        =   1
      Top             =   975
      Width           =   1740
   End
   Begin VB.TextBox txtName 
      Height          =   360
      Left            =   375
      TabIndex        =   0
      Top             =   300
      Width           =   1740
   End
   Begin VB.Label lblView 
      Height          =   390
      Left            =   75
      TabIndex        =   5
      Top             =   1875
      Width           =   2565
   End
   Begin VB.Label Label2 
      Caption         =   "Telephone"
      Height          =   240
      Left            =   375
      TabIndex        =   3
      Top             =   750
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   165
      Left            =   375
      TabIndex        =   2
      Top             =   75
      Width           =   1740
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
    Dim contact As New cContact
    
    contact.setValue "name", txtName.Text
    contact.setValue "phone", txtPhone.Text
    
    lblView.Caption = "Contact Name:" & contact.getValue("name") & " Phone:" & contact.getValue("phone")
    
    Dim people As New cPeople
    
    people.setValue "nome", txtName.Text
    people.setValue "telefone", txtPhone.Text
        
    MsgBox people("nome")
    
    people.setValue "codigo", 3
    
    If Not people.getById.EOF Then
        MsgBox people("nome") & ";" & people("telefone") & ";" & people("endereco")
    End If
    
    Debug.Print people.getById.getString(, , "|", Chr(13))
    Debug.Print people.getString("|")
    
    
    Dim produto As New clsClassDefault
    produto.buildClass "cad_produtos", "codigo", Conn
    produto.setValue "codigo", 4
    produto.getById
    
    Debug.Print produto.getString
    
    Dim rs As ADODB.Recordset
    
    Set rs = produto.getListAll(10)
    
    
    produto.getListFilter "descricao LIKE '%a%'"
End Sub

Private Sub Form_Load()
    Conn.Open "dsn=he_varejo"
End Sub
