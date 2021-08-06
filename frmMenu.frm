VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oficina 1.0"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8655
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btComanda 
      Caption         =   "Comanda"
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btServiços 
      Caption         =   "Serviços"
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btClientes 
      Caption         =   "Clientes"
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   6015
      Left            =   0
      Picture         =   "frmMenu.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "B. São Benedito - Uberaba MG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   3000
      Width           =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rua Exemplo do teste, 999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   2760
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Linha Reta Serviços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3960
      TabIndex        =   1
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Menu menarquivos 
      Caption         =   "&Arquivos"
      Begin VB.Menu mensair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mensobre 
      Caption         =   "Sobre..."
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btClientes_Click()
    frmclientes.Show 1
End Sub

Private Sub btComanda_Click()
    frmComanda.Show 1
End Sub

Private Sub btSair_Click()
    Dim Resp As Integer
    Resp = MsgBox("Fechar aplicativo ?", 4 + 32 + 256, "Confirmação")
    If Resp = 6 Then
        Unload Me
        End
    End If
End Sub

Private Sub btServiços_Click()
    frmservicos.Show 1
End Sub

Private Sub Form_Load()
    ORIGEM = App.Path & "\"
    DESTINO = App.Path & "\"
End Sub

Private Sub mensair_Click()
 Call btSair_Click
End Sub

Private Sub mensobre_Click()
    frmSobre.Show 1
End Sub
