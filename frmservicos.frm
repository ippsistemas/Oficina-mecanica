VERSION 5.00
Begin VB.Form frmservicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Serviços"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   ControlBox      =   0   'False
   Icon            =   "frmservicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3360
      TabIndex        =   14
      Top             =   3480
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
      Left            =   2880
      TabIndex        =   13
      Top             =   3480
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
      Left            =   2400
      TabIndex        =   12
      Top             =   3480
      Width           =   495
   End
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
      Left            =   1920
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton btIncluir 
         Caption         =   "Incluir"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton btSalvar 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1275
         TabIndex        =   9
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton btExcluir 
         Caption         =   "Excluir"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2445
         TabIndex        =   8
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton btFechar 
         Caption         =   "Fechar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4680
         TabIndex        =   7
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.TextBox CPTEXTO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox CPTEXTO 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   5535
   End
   Begin VB.TextBox CPTEXTO 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   975
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
      Left            =   1920
      TabIndex        =   15
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Preço Unitário"
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
      TabIndex        =   5
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição do Serviço"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   600
   End
End
Attribute VB_Name = "frmservicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btExcluir_Click()
    'EXCLUIR UM REGISTRO
    Dim Resp As Integer
    
    Resp = MsgBox("Excluir registro ?", 4 + 32 + 256, "Confirmação")
    If Resp = 6 Then
            On Error Resume Next
            Open ORIGEM & "DADOS.DAT" For Random As #20 Len = Len(REG)
            RSERVICOS.DELETADO = "S"
            Put #1, POS, RSERVICOS
            Close #1
            CPTEXTO(0).Text = ""
            CPTEXTO(1).Text = ""
            CPTEXTO(2).Text = ""
            CPTEXTO(0).SetFocus
            MsgBox "Registro excluído!", vbInformation, "Oficina!"
            Call COMPACTAR
            Open ORIGEM & "SERVICOS.DAT" For Random As #1 Len = Len(RSERVICOS)
            Call SHOWSERVICO
            
    End If
End Sub

Private Sub btFechar_Click()
    'Fechar arquivo
     Close #1
     Unload Me
End Sub

Private Sub btIncluir_Click()
    frmservicos.Caption = "Cadastro de Serviços [ Incluindo ]"
    
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
    CPTEXTO(0).SetFocus
    
    'GERAR PROXIMO
    COMP = Len(RSERVICOS)
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
    Open ORIGEM & "servicos.dat" For Random As #1 Len = Len(RSERVICOS)
    Call SHOWSERVICO
End Sub

Private Sub SHOWSERVICO()
    COMP = Len(RSERVICOS)
    TAM = LOF(1) / COMP
    If POS > TAM Then POS = TAM
    If TAM <> 0 Then

            
            Get #1, POS, RSERVICOS
            
            CPTEXTO(0).Text = RSERVICOS.CODSERVICO
            CPTEXTO(1).Text = RSERVICOS.NOMESERVICO
            CPTEXTO(2).Text = Format(RSERVICOS.Preco, "##,##0.00")
            
            Call NAVEGANDO
    Else
            frmservicos.Caption = "Cadastro de Serviços [ Arquivo vazio ]"
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
        RSERVICOS.CODSERVICO = CPTEXTO(0).Text
        RSERVICOS.NOMESERVICO = CPTEXTO(1).Text
        RSERVICOS.Preco = CPTEXTO(2).Text
        RSERVICOS.DELETADO = "N"
        
        If POS = 0 Then POS = POS + 1
        Put #1, POS, RSERVICOS
        Call NAVEGANDO
        'MsgBox "Gravado com sucesso!", vbInformation, "Oficina!"
        
End Sub

Private Sub NAVEGANDO()
    COMP = Len(RSERVICOS)
    TAM = LOF(1) / COMP
    
    LPOS.Caption = Str(POS) & "/" & Str(TAM)
    frmservicos.Caption = "Cadastro de Serviços [ Navegando ]"
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
    Call SHOWSERVICO

End Sub


Private Sub BTPRIMEIRO_Click()
    POS = 1
    Call SHOWSERVICO

End Sub

Private Sub BTPROXIMO_Click()
    On Error Resume Next
    POS = POS + 1
    If POS > TAM Then POS = TAM
    Call SHOWSERVICO

End Sub

Private Sub BTANTERIOR_Click()
    On Error Resume Next
    POS = POS - 1
    If POS > 0 Then Call SHOWSERVICO

End Sub
Private Sub COMPACTAR()
            Dim i As Integer
            
            'COMPACTANDO DADOS.DAT
            If EXISTE(DESTINO & "DADOS.BAK") Then Kill (DESTINO & "DADOS.BAK")
            COMP = Len(RSERVICOS)
            On Error Resume Next
            Open ORIGEM & "SERVICOS.DAT" For Random As #1 Len = COMP
            Open DESTINO & "DADOS.BAK" For Random As #2 Len = COMP
            TAM = LOF(1) / COMP
            For i = 1 To TAM
                Get #1, , RSERVICOS
                If RSERVICOS.DELETADO <> "S" Then
                    Put #2, , RSERVICOS
                End If
            Next
            Close #1
            Close #2
            Kill (ORIGEM & "SERVICOS.DAT")
            Name DESTINO & "DADOS.BAK" As ORIGEM & "SERVICOS.DAT"

End Sub
