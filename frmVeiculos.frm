VERSION 5.00
Begin VB.Form frmComanda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão Comanda"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "frmVeiculos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton btnAcao 
         Caption         =   "Incluir"
         Height          =   495
         Index           =   9
         Left            =   360
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton btnAcao 
         Caption         =   "Alterar"
         Height          =   495
         Index           =   8
         Left            =   1500
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton btnAcao 
         Caption         =   "Salvar"
         Height          =   495
         Index           =   7
         Left            =   2640
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton btnAcao 
         Caption         =   "Excluir"
         Height          =   495
         Index           =   6
         Left            =   3780
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton btnAcao 
         Caption         =   "Fechar"
         Height          =   495
         Index           =   5
         Left            =   4920
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   1230
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   6255
   End
   Begin VB.TextBox CPTEXTO 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Todos veículos deste Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3015
      Width           =   2070
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Código/CPF"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Modelo"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cor"
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ano"
      Height          =   195
      Left            =   5280
      TabIndex        =   8
      Top             =   1440
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Placa"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Proprietario"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   1485
   End
End
Attribute VB_Name = "frmComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAcao_Click(Index As Integer)
    Unload Me
End Sub
