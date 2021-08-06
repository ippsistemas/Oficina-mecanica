VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComanda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão Comanda"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmComanda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic1 
      BackColor       =   &H00C0FFFF&
      Height          =   6735
      Left            =   6480
      ScaleHeight     =   6675
      ScaleWidth      =   4995
      TabIndex        =   32
      Top             =   120
      Width           =   5055
   End
   Begin VB.TextBox Numero 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1240
      Width           =   1695
   End
   Begin VB.CommandButton btAdd 
      Height          =   375
      Left            =   5400
      Picture         =   "frmComanda.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Celular 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   2340
      Width           =   1935
   End
   Begin VB.ComboBox Descricao 
      Height          =   315
      Left            =   720
      TabIndex        =   8
      Top             =   3750
      Width           =   3495
   End
   Begin VB.ComboBox NomeCliente 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1620
      Width           =   5175
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   4320
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   5670
      Width           =   2055
   End
   Begin VB.TextBox Qtd 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   3750
      Width           =   495
   End
   Begin VB.TextBox Preco 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Text            =   "0,00"
      Top             =   3750
      Width           =   975
   End
   Begin VB.TextBox TelCliente 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   2340
      Width           =   1935
   End
   Begin VB.TextBox EndCliente 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   1980
      Width           =   5175
   End
   Begin VB.TextBox DataServico 
      Height          =   315
      Left            =   5160
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1240
      Width           =   1215
   End
   Begin VB.TextBox Bairro 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Text            =   "B. Exemplo -UBERABA MG- isaiaspereirapinto@gmail.com"
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox Endereco 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Text            =   "RUA ANTONIO 999"
      Top             =   480
      Width           =   6255
   End
   Begin VB.TextBox Titulo 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Text            =   "LINHA RETA SERVIÇOS"
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton BTFECHAR 
      Caption         =   "Fechar"
      Height          =   735
      Left            =   4440
      Picture         =   "frmComanda.frx":0106
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton btLimpa 
      Caption         =   "Limpar campos"
      Height          =   735
      Left            =   600
      Picture         =   "frmComanda.frx":0410
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton btImprimir 
      Caption         =   "Gerar"
      Height          =   735
      Left            =   2520
      Picture         =   "frmComanda.frx":071A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Km 
      Height          =   315
      Left            =   5280
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2700
      Width           =   1095
   End
   Begin VB.TextBox Placa 
      Height          =   315
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2700
      Width           =   1095
   End
   Begin VB.TextBox Veiculo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2700
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   38
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "PREÇO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVIÇO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   36
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "QTD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1280
      Width           =   735
   End
   Begin VB.Label lblCelular 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Celular"
      Height          =   195
      Left            =   3600
      TabIndex        =   31
      Top             =   2430
      Width           =   480
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Geral"
      Height          =   195
      Left            =   3360
      TabIndex        =   30
      Top             =   5760
      Width           =   780
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd."
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3525
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço"
      Height          =   195
      Left            =   4320
      TabIndex        =   27
      Top             =   3525
      Width           =   420
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do Serviço"
      Height          =   195
      Left            =   720
      TabIndex        =   26
      Top             =   3525
      Width           =   1530
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1980
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   4680
      TabIndex        =   23
      Top             =   1275
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serviços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      Height          =   195
      Left            =   4920
      TabIndex        =   15
      Top             =   2760
      Width           =   225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa"
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   2760
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Veículo"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1650
      Width           =   480
   End
End
Attribute VB_Name = "frmComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTALCOMANDA As Currency
Dim liberado As Boolean
Dim l, c As Double


Private Sub btAdd_Click()
  If (Qtd.Text <> "" And Descricao.Text <> "") Then
    linha = linha + 200
    coluna = 100

    'GRAVA CAMPOS COMANDA E CAMPOS DETALHESCOMANDA
    RCOMANDA.DTCOMANDA = DataServico.Text
    RCOMANDA.CLIENTE = NomeCliente.Text
    RCOMANDA.Veiculo = Veiculo.Text
    RCOMANDA.Placa = Placa.Text
    RCOMANDA.Km = Km.Text
        
    'DET DA COMANDA
    RDTCOMANDA.NCOMANDA = Numero.Text
    RDTCOMANDA.DELETADO = "N"
    RDTCOMANDA.QUANTIDADE = Qtd.Text
    RDTCOMANDA.NOMESERVICO = Descricao.Text
    RDTCOMANDA.Preco = Preco.Text
    RDTCOMANDA.SUBTOTAL = Format(RDTCOMANDA.Preco * RDTCOMANDA.QUANTIDADE, "##,##0.00")
    Put #6, DPOS, RDTCOMANDA
    DPOS = DPOS + 1
    
    frmComanda.ForeColor = &HC00000
    frmComanda.CurrentX = coluna
    frmComanda.CurrentY = linha
    Print Format(RDTCOMANDA.QUANTIDADE, "0000")
    
    frmComanda.CurrentX = coluna + 600
    frmComanda.CurrentY = linha
    Print Left(RDTCOMANDA.NOMESERVICO, 30)
    
    frmComanda.CurrentX = coluna + 4200
    frmComanda.CurrentY = linha
    Print Format(RDTCOMANDA.Preco, "##,##0.00")
    
    frmComanda.CurrentX = coluna + 5600
    frmComanda.CurrentY = linha
    Print Format(RDTCOMANDA.SUBTOTAL, "##,##0.00")
    
    TOTALCOMANDA = TOTALCOMANDA + RDTCOMANDA.SUBTOTAL
    Total.Text = Format(TOTALCOMANDA, "##,##0.00")
    RCOMANDA.Total = TOTALCOMANDA
    
    Qtd.SetFocus
  End If
End Sub

Private Sub btFechar_Click()
    Unload Me
End Sub

Private Sub btImprimir_Click()
 
 Dim PaginaInicial, Paginafinal, numerodecopias, i As Integer
 CommonDialog1.CancelError = True

 frmComanda.Width = 10530
 
 'FECHO A DETALHE COMANDA
 If POS = 0 Then POS = POS + 1
 Put #5, POS, RCOMANDA
 Close #5
  
 frmComanda.MousePointer = 11
 btImprimir.Enabled = False
           
 On Error GoTo TrataErro
 CommonDialog1.ShowPrinter
 
 'Captura os valores definidos pelo usuário na janela
 PaginaInicial = CommonDialog1.FromPage
 Paginafinal = CommonDialog1.ToPage
 numerodecopias = CommonDialog1.Copies
 
 'CONFIRMANDO CONFIGURAÇÃO IMPRESSORA
 Call ConfiguraImp
'IMPRIME PICTURE
 Call ImprimePic
 
 'IMPRIME NA IMPRESSORA
 Printer.PaperSize = vbPRPSA4 '(define o tamanho do papel para: A4 , 210 x 297 mm)
 Printer.Zoom = 50  '(Define o zoom em 50% do tamanho original)
 Printer.Orientation = vbPRORPortrait
 
 For i = 1 To numerodecopias
    Call ImprimeImp
    If numerodecopias > 1 Then Printer.NewPage
 Next i

 frmComanda.MousePointer = 0
 btImprimir.Enabled = False

TrataErro:
    Exit Sub

End Sub

Private Sub btLimpa_Click()
    DataServico.Text = Date
    NomeCliente.Text = ""
    EndCliente.Text = ""
    TelCliente.Text = ""
    Veiculo.Text = ""
    Placa.Text = ""
    Km.Text = ""
    Qtd.Text = ""
    Descricao.Text = ""
    Preco.Text = "0,00"
    'Lista.Clear
    Total.Text = "0,00"
    linha = 100
    coluna = 4500
    
End Sub

Private Sub Descricao_LostFocus()
    Dim i As Integer
    
    Open ORIGEM & "SERVICOS.DAT" For Random As #20 Len = Len(RSERVICOS)
    COMP = Len(RSERVICOS)
    TAM = LOF(20) / COMP
    If TAM <> 0 Then
       For i = 1 To TAM
           Get #20, i, RSERVICOS
           'PREENCHER CAMPOS
           If Trim(Descricao.Text) = Trim(RSERVICOS.NOMESERVICO) Then
              Preco.Text = Format(RSERVICOS.Preco, "##,##0.00")
           End If
       Next i
    End If
    Close #20
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = 13 Then
    'Se o tipo do controle ativo for TextBox
    'If TypeOf Screen.ActiveControl Is TextBox Then
     'Simula o pressionamento da tecla TAB
      SendKeys "{tab}"
      'A linha a seguir evita ouvir um bip
      KeyAscii = 0
    'End If
  End If
End Sub

Private Sub Form_Load()
    Call novacomanda
    Call btLimpa_Click
    Call LERCLIENTES
    Call LERSERVICOS
    linha = 4200
    coluna = 100
    liberado = False

    
End Sub

Private Sub NomeCliente_LostFocus()
    Dim i As Integer
    
    Open ORIGEM & "CLIENTES.DAT" For Random As #20 Len = Len(RCLIENTES)
    COMP = Len(RCLIENTES)
    TAM = LOF(20) / COMP
    If TAM <> 0 Then
       For i = 1 To TAM
           Get #20, i, RCLIENTES
           'PREENCHER CAMPOS
           If Trim(NomeCliente.Text) = Trim(RCLIENTES.NOME) Then
              EndCliente.Text = RCLIENTES.Endereco
              TelCliente.Text = RCLIENTES.TELEFONE
              Celular.Text = RCLIENTES.Celular
     
           End If
       Next i
    End If
    Close #20

End Sub


Private Sub LERCLIENTES()
    Dim i As Integer
    
    Open ORIGEM & "CLIENTES.DAT" For Random As #20 Len = Len(RCLIENTES)
    COMP = Len(RCLIENTES)
    TAM = LOF(20) / COMP
    If TAM <> 0 Then
       For i = 1 To TAM
           Get #20, i, RCLIENTES
           NomeCliente.AddItem Trim(RCLIENTES.NOME)
       Next i
    End If
    Close #20

End Sub

Private Sub LERSERVICOS()
    Dim i As Integer
    
    Open ORIGEM & "SeRVICOS.DAT" For Random As #20 Len = Len(RSERVICOS)
    COMP = Len(RSERVICOS)
    TAM = LOF(20) / COMP
    If TAM <> 0 Then
       For i = 1 To TAM
           Get #20, i, RSERVICOS
           Descricao.AddItem Trim(RSERVICOS.NOMESERVICO)
       Next i
    End If
    Close #20

End Sub

Private Sub updtotal()
   Dim valor As Currency
   valor = Val(Total.Text)
   'valor = valor + Val(Subtotal.Text)
   Total.Text = Format(valor, "##,##0.00")
End Sub


Private Sub Pic1_Click()
    'imprimir na impressora
    liberado = False
    Unload Me
End Sub

Private Sub Qtd_GotFocus()
    Qtd.Text = ""
End Sub

Private Sub Qtd_LostFocus()
    If Qtd.Text = "" Then Qtd.Text = 1
    
End Sub

Private Sub novacomanda()
    'GERAR UM NOVO NUMERO
    On Error Resume Next
    Open ORIGEM & "GERADOS.BIN" For Binary As #13  'ABRIR ARQUIVO
    Get #13, 1, n         'LER DADO
    n = n + 1            'INCREMENTAR
    Put #13, 1, n        'GRAVAR NUMERO GRAVEI OK
    Close #13             'FECHA O ARQUIVO SEM GRAVAR O NUMERO
    NUMEROGERADO = Format(n, "000000")
    Numero.Text = NUMEROGERADO
    
    'NOVA COMANDA
    On Error Resume Next
    Open ORIGEM & "COMANDA.DAT" For Random As #5 Len = Len(RCOMANDA)
    COMP = Len(RCOMANDA)
    TAM = LOF(5) / COMP
    POS = TAM + 1
    RCOMANDA.NCOMANDA = Numero.Text
    RCOMANDA.DELETADO = "N"
        
    'DET DA COMANDA
    On Error Resume Next
    Open ORIGEM & "DETCOMANDA.DAT" For Random As #6 Len = Len(RDTCOMANDA)
    DCOMP = Len(RDTCOMANDA)
    DTAM = LOF(6) / DCOMP
    DPOS = DTAM + 1
    
    TOTALCOMANDA = 0
End Sub



Private Sub ConfiguraImp()
    Pic1.ScaleMode = 7
    Pic1.Width = 567 * 7     '(22 centimetros)
    Pic1.Height = 567 * 12    '(14 centimetros   )
    Pic1.Font = "Tahoma"
    Pic1.FontSize = 8
    Pic1.Font.Bold = False
    Pic1.Font.Italic = False

End Sub
Private Sub ImprimePic()
     Pic1.Cls
     Pic1.CurrentX = 1.8
     Pic1.CurrentY = 0
     Pic1.Print Titulo.Text
     Pic1.CurrentX = 0.1
     Pic1.CurrentY = 0.1
     Pic1.Print String(42, "_")
     
     Pic1.CurrentX = 0.1
     Pic1.CurrentY = 0.5
     Pic1.Print Endereco.Text
     
     Pic1.CurrentX = 0.1
     Pic1.CurrentY = 0.8
     Pic1.Print Bairro.Text
    
     Pic1.CurrentX = 0.1
     Pic1.CurrentY = 0.9
     Pic1.Print String(42, "_")
     
     Pic1.CurrentX = 0.1
     Pic1.CurrentY = 1.3
     Pic1.Print DataServico.Text + "   Hora:" + Format(Time, "HH:MM")
 
 Pic1.CurrentX = 5.1
 Pic1.CurrentY = 1.3
 Pic1.Print "COD:" + Numero
 
 Pic1.CurrentX = 0.1
 Pic1.CurrentY = 1.4
 Pic1.Print String(42, "_")
 
 Pic1.CurrentX = 0.1
 Pic1.CurrentY = 1.8
 Pic1.Print NomeCliente.Text
 Pic1.CurrentX = 4.7
 Pic1.CurrentY = 1.8
 Pic1.Print TelCliente.Text
 Pic1.CurrentX = 0.1
 Pic1.CurrentY = 2.1
 Pic1.Print EndCliente.Text

 Pic1.CurrentX = 0.1
 Pic1.CurrentY = 2.4
 Pic1.Print "Veíc.:" + Veiculo.Text
 Pic1.CurrentX = 3
 Pic1.CurrentY = 2.4
 Pic1.Print "Pl:" + Placa.Text
 Pic1.CurrentX = 4.9
 Pic1.CurrentY = 2.4
 Pic1.Print "Km:" + Km.Text
 Pic1.CurrentX = 0.1
 Pic1.CurrentY = 2.6
 Pic1.Print String(42, "_")
 
        'DET DA COMANDA
        On Error Resume Next
        Open ORIGEM & "DETCOMANDA.DAT" For Random As #6 Len = Len(RDTCOMANDA)
        DCOMP = Len(RDTCOMANDA)
        DTAM = LOF(6) / DCOMP
        
        c = 0.1
        l = 3
        
        If DTAM > 0 Then
          For i = 1 To DTAM
              Get #6, i, RDTCOMANDA
              If RDTCOMANDA.NCOMANDA = Val(Numero.Text) Then
                        Pic1.CurrentX = c
                        Pic1.CurrentY = l
                        Pic1.Print Format(RDTCOMANDA.QUANTIDADE, "0000")
                        
                        Pic1.CurrentX = c + 0.7
                        Pic1.CurrentY = l
                        Pic1.Print Left(RDTCOMANDA.NOMESERVICO, 20)
                        
                        Pic1.CurrentX = c + 4.2
                        Pic1.CurrentY = l
                        Pic1.Print Format(RDTCOMANDA.Preco, "##,##0.00")
                        
                        Pic1.CurrentX = c + 5.7
                        Pic1.CurrentY = l
                        Pic1.Print Format(RDTCOMANDA.SUBTOTAL, "##,##0.00")

                        l = l + 0.4
                        
              End If
          Next i
           
          Pic1.CurrentX = 0.1
          Pic1.CurrentY = l
          Pic1.Print String(42, "_")
          
          Pic1.CurrentX = c + 5.7
          Pic1.CurrentY = l + 0.4
          Pic1.Print Format(Total.Text, "##,##0.00")
            

        End If
        Close #6

End Sub

Private Sub ImprimeImp()
 'Configura impressora
 Printer.ScaleMode = 7
 Printer.Width = 567 * 26 '(26 centimetros)
 Printer.Height = 567 * 10.2 '(10.2 centimetros)
 Printer.Font = "Courier New"
 Printer.FontSize = 10
 Printer.Font.Bold = False
 Printer.Font.Italic = False
 Printer.EndDoc
 
 'CONFIRMANDO CONFIGURAÇÃO
 Printer.ScaleMode = 7
 Printer.Width = 567 * 26       '(26 centimetros)
 Printer.Height = 567 * 10.2 '(10.2 centimetros)
 Printer.Font = "Courier New"
 Printer.FontSize = 10
 Printer.Font.Bold = False
 Printer.Font.Italic = False
 Printer.EndDoc
 
'INICIAR A IMPRESSÃO
 On Error Resume Next
     
 Printer.CurrentX = 2
 Printer.CurrentY = 0.3
 Printer.Print Titulo.Text
 Printer.CurrentX = 8
 Printer.CurrentY = 0.3
 Printer.Print String(42, "_")
 
 Printer.CurrentX = 14
 Printer.CurrentY = 0.3
 Printer.Print Endereco.Text
 
 Printer.CurrentX = 15
 Printer.CurrentY = 0.3
 Printer.Print Bairro.Text

 Printer.CurrentX = 16
 Printer.CurrentY = 0.3
 Printer.Print String(42, "_")
 
 Printer.CurrentX = 17
 Printer.CurrentY = 0.3
 Printer.Print DataServico.Text + "   Hora:" + Format(Time, "HH:MM")
 
 Printer.CurrentX = 18
 Printer.CurrentY = 0.3
 Printer.Print "COD:" + Numero
 
 Printer.CurrentX = 19
 Printer.CurrentY = 0.3
 Printer.Print String(42, "_")
 
 Printer.CurrentX = 20
 Printer.CurrentY = 0.3
 Printer.Print NomeCliente.Text
 
 Printer.CurrentX = 4.7
 Printer.CurrentY = 1.8
 Printer.Print TelCliente.Text
 
 Printer.CurrentX = 0.1
 Printer.CurrentY = 2.1
 Printer.Print EndCliente.Text

 Printer.CurrentX = 0.1
 Printer.CurrentY = 2.4
 Printer.Print "Veíc.:" + Veiculo.Text
 Printer.CurrentX = 3
 Printer.CurrentY = 2.4
 Printer.Print "Pl:" + Placa.Text
 Printer.CurrentX = 4.9
 Printer.CurrentY = 2.4
 Printer.Print "Km:" + Km.Text
 Printer.CurrentX = 0.1
 Printer.CurrentY = 2.6
 Printer.Print String(42, "_")
 
'DET DA COMANDA
'On Error Resume Next
'Open ORIGEM & "DETCOMANDA.DAT" For Random As #6 Len = Len(RDTCOMANDA)
'DCOMP = Len(RDTCOMANDA)
'DTAM = LOF(6) / DCOMP

'c = 0.1
'l = 3

'If DTAM > 0 Then
'  For i = 1 To DTAM
'      Get #6, i, RDTCOMANDA
'      If RDTCOMANDA.NCOMANDA = Val(Numero.Text) Then
'                Printer.CurrentX = c
'                Printer.CurrentY = l
'                Printer.Print Format(RDTCOMANDA.QUANTIDADE, "0000")
'
'                Printer.CurrentX = c + 0.7
'                Printer.CurrentY = l
'                Printer.Print Left(RDTCOMANDA.NOMESERVICO, 20)
'
'                Printer.CurrentX = c + 4.2
'                Printer.CurrentY = l
'                Printer.Print Format(RDTCOMANDA.Preco, "##,##0.00")
'
'                Printer.CurrentX = c + 5.7
'                Printer.CurrentY = l
'                Printer.Print Format(RDTCOMANDA.SUBTOTAL, "##,##0.00")
'
'                l = l + 0.4
'
'      End If
'  Next i
   
'  Printer.CurrentX = 0.1
'  Printer.CurrentY = l
'  Printer.Print String(42, "_")
'
'  Printer.CurrentX = c + 5.7
'  Printer.CurrentY = l + 0.4
'  Printer.Print Format(Total.Text, "##,##0.00")
    

'End If

Close #6

Printer.EndDoc

End Sub

