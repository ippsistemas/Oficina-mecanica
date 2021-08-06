Attribute VB_Name = "Modulo"
Option Explicit
Public ORIGEM, DESTINO As String
Public POS, DPOS, linha, coluna As Integer
Public COMP, DCOMP As Long
Public TAM, DTAM As Long

'* * * * * * * * * * * * * * * * * * * * * * * * * * * *

Type TBCLIENTES
    CODCLIENTE As String * 14
    NOME As String * 50
    Endereco As String * 50
    Bairro As String * 30
    CIDADE As String * 30
    ESTADO As String * 2
    CEP As String * 9
    TELEFONE As String * 14
    Celular As String * 14
    DELETADO As String * 1
End Type
Public RCLIENTES As TBCLIENTES

'* * * * * * * * * * * * * * * * * * * * * * * * * * * *

Type TBSERVICOS
   CODSERVICO As String * 4
   NOMESERVICO As String * 50
   Preco As Currency
   DELETADO As String * 1
End Type
Public RSERVICOS As TBSERVICOS

'* * * * * * * * * * * * * * * * * * * * * * * * * * * *

Type TBCOMANDA
    NCOMANDA As Long
    DTCOMANDA As String * 10
    CLIENTE As String * 50
    Veiculo As String * 15
    Placa As String * 8
    Km As String * 10
    Total As Currency
    DELETADO As String * 1
End Type
Public RCOMANDA As TBCOMANDA

'* * * * * * * * * * * * * * * * * * * * * * * * * * * *

Type TBDTCOMANDA
    NCOMANDA As Long
    QUANTIDADE As Integer
    NOMESERVICO As String * 50
    Preco As Currency
    SUBTOTAL As Currency
    DELETADO As String * 1
End Type
Public RDTCOMANDA As TBDTCOMANDA

'* * * * * * * * * * * * * * * * * * * * * * * * * * * *

'TESTA SE ARQUIVO EXISTE
Public Function EXISTE(NomeArq As String) As Integer
    EXISTE = Len(Dir$(NomeArq$)) > 0
End Function
