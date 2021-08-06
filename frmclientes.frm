VERSION 5.00
Begin VB.Form frmclientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   5790
   ClientLeft      =   345
   ClientTop       =   495
   ClientWidth     =   5850
   ControlBox      =   0   'False
   Icon            =   "frmclientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BTPRIMEIRO 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   24
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton BTANTERIOR 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   23
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton BTPROXIMO 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton BTULTIMO 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   21
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton BTFECHAR 
         Caption         =   "Fechar"
         Height          =   495
         Left            =   4680
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton BTEXCLUIR 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   2340
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton BTSALVAR 
         Caption         =   "Salvar"
         Height          =   495
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton BTINCLUIR 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox CPTEXTO 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   5535
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Label LPOS 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Celular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3840
      TabIndex        =   20
      Top             =   4080
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1920
      TabIndex        =   19
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código/CPF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   825
   End
End
Attribute VB_Name = "frmclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btAlterar_Click()
    frmclientes.Caption = "Cadastro de Clientes [ Editando ]"

    BTSALVAR.Enabled = True
    BTALTERAR.Enabled = False
    BTINCLUIR.Enabled = False
    BTEXCLUIR.Enabled = False
    BTFECHAR.Enabled = True

End Sub

Private Sub btExcluir_Click()
    'EXCLUIR UM REGISTRO
    Dim Resp As Integer
    
    Resp = MsgBox("Excluir registro ?", 4 + 32 + 256, "Confirmação")
    If Resp = 6 Then
            On Error Resume Next
            Open ORIGEM & "DADOS.DAT" For Random As #20 Len = Len(REG)
            RCLIENTES.DELETADO = "S"
            Put #1, POS, RSERVICOS
            Close #1
            CPTEXTO(0).Text = ""
            CPTEXTO(1).Text = ""
            CPTEXTO(2).Text = ""
            CPTEXTO(3).Text = ""
            CPTEXTO(4).Text = ""
            CPTEXTO(5).Text = ""
            CPTEXTO(6).Text = ""
            CPTEXTO(7).Text = ""
    
    CPTEXTO(0).SetFocus
            MsgBox "Registro excluído!", vbInformation, "Oficina!"
            Call COMPACTAR
            Open ORIGEM & "CLIENTES.DAT" For Random As #1 Len = Len(RCLIENTES)
            Call SHOWCLIENTE
            
    End If
End Sub

Private Sub btFechar_Click()
    'Fechar arquivo
     Close #1
     Unload Me
End Sub

Private Sub btIncluir_Click()
    frmclientes.Caption = "Cadastro de Clientes [ Incluindo ]"
    
    'BTALTERAR.Enabled = False
    BTEXCLUIR.Enabled = False
    BTSALVAR.Enabled = True
    BTINCLUIR.Enabled = False
    BTFECHAR.Enabled = False
    
    BTPRIMEIRO.Enabled = False
    BTANTERIOR.Enabled = False
    BTPROXIMO.Enabled = False
    BTULTIMO.Enabled = False
    
    CPTEXTO(0).Text = ""
    CPTEXTO(1).Text = ""
    CPTEXTO(2).Text = ""
    CPTEXTO(3).Text = ""
    CPTEXTO(4).Text = ""
    CPTEXTO(5).Text = ""
    CPTEXTO(6).Text = ""
    CPTEXTO(7).Text = ""
    
    CPTEXTO(0).SetFocus
    'GERAR PROXIMO
    COMP = Len(RCLIENTES)
    TAM = LOF(1) / COMP
    POS = TAM + 1
    LPOS.Caption = Str(POS) & "/" & Str(TAM + 1)
End Sub

Private Sub btSalvar_Click()
    Call gravando
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
POS = 1
    'abrir arquivo
    On Error Resume Next
    Open ORIGEM & "CLIENTES.DAT" For Random As #1 Len = Len(RCLIENTES)
    Call SHOWCLIENTE
End Sub

Private Sub SHOWCLIENTE()
    COMP = Len(RCLIENTES)
    TAM = LOF(1) / COMP
    If POS > TAM Then POS = TAM
    If TAM <> 0 Then

           
            Get #1, POS, RCLIENTES
            
            CPTEXTO(0).Text = RCLIENTES.CODCLIENTE
            CPTEXTO(1).Text = RCLIENTES.NOME
            CPTEXTO(2).Text = RCLIENTES.Endereco
            CPTEXTO(3).Text = RCLIENTES.Bairro
            CPTEXTO(4).Text = RCLIENTES.CIDADE
            CPTEXTO(5).Text = RCLIENTES.CEP
            CPTEXTO(6).Text = RCLIENTES.TELEFONE
            CPTEXTO(7).Text = RCLIENTES.Celular
            
            Call NAVEGANDO
    Else
            frmclientes.Caption = "Cadastro de Clientes [ Arquivo vazio ]"
            BTINCLUIR.Enabled = False
            BTSALVAR.Enabled = True
            BTALTERAR.Enabled = False
            BTEXCLUIR.Enabled = False
            BTFECHAR.Enabled = True
            
            BTPRIMEIRO.Enabled = False
            BTANTERIOR.Enabled = False
            BTPROXIMO.Enabled = False
            BTULTIMO.Enabled = False
            
    End If
            LPOS.Caption = Str(POS) & "/" & Str(TAM)

End Sub
Private Sub gravando()
            RCLIENTES.CODCLIENTE = CPTEXTO(0).Text
            RCLIENTES.NOME = CPTEXTO(1).Text
            RCLIENTES.Endereco = CPTEXTO(2).Text
            RCLIENTES.Bairro = CPTEXTO(3).Text
            RCLIENTES.CIDADE = CPTEXTO(4).Text
            RCLIENTES.CEP = CPTEXTO(5).Text
            RCLIENTES.TELEFONE = CPTEXTO(6).Text
            RCLIENTES.Celular = CPTEXTO(7).Text
            RCLIENTES.DELETADO = "N"
            
        If POS = 0 Then POS = POS + 1
        Put #1, POS, RCLIENTES
        Call NAVEGANDO
        'MsgBox "Gravado com sucesso!", vbInformation, "Oficina!"
        
End Sub

Private Sub NAVEGANDO()
    COMP = Len(RCLIENTES)
    TAM = LOF(1) / COMP
    LPOS.Caption = Str(POS) & "/" & Str(TAM)
    frmclientes.Caption = "Cadastro de Clientes [ Navegando ]"
    BTINCLUIR.Enabled = True
    BTSALVAR.Enabled = True
    'BTALTERAR.Enabled = True
    BTEXCLUIR.Enabled = True
    BTFECHAR.Enabled = True
    
    BTPRIMEIRO.Enabled = True
    BTANTERIOR.Enabled = True
    BTPROXIMO.Enabled = True
    BTULTIMO.Enabled = True
    
End Sub
Private Sub BTULTIMO_Click()
    POS = TAM
    Call SHOWCLIENTE

End Sub


Private Sub BTPRIMEIRO_Click()
    POS = 1
    Call SHOWCLIENTE

End Sub

Private Sub BTPROXIMO_Click()
    On Error Resume Next
    POS = POS + 1
    If POS > TAM Then POS = TAM
    Call SHOWCLIENTE

End Sub

Private Sub BTANTERIOR_Click()
    On Error Resume Next
    POS = POS - 1
    If POS > 0 Then Call SHOWCLIENTE

End Sub
Private Sub COMPACTAR()
            Dim i As Integer
            
            'COMPACTANDO DADOS.DAT
            If EXISTE(DESTINO & "DADOS.BAK") Then Kill (DESTINO & "DADOS.BAK")
            COMP = Len(RCLIENTES)
            On Error Resume Next
            Open ORIGEM & "CLINTES.DAT" For Random As #1 Len = COMP
            Open DESTINO & "DADOS.BAK" For Random As #2 Len = COMP
            TAM = LOF(1) / COMP
            For i = 1 To TAM
                Get #1, , RCLIENTES
                If RCLIENTES.DELETADO <> "S" Then
                    Put #2, , RCLIENTES
                End If
            Next
            Close #1
            Close #2
            Kill (ORIGEM & "CLIENTES.DAT")
            Name DESTINO & "DADOS.BAK" As ORIGEM & "CLIENTES.DAT"

End Sub

