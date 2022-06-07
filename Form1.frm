VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CSS-Sisemas"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Teste Impressão"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label 
      Caption         =   "Teste de Impressao Etiqueta PPLB "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DorvaOS 12345, "Tede de Cliente", "243,56"
End Sub


Sub DorvaOS(numeroOS As Long, Nome_Cliente As String, WPRECO As String)
Dim i As Integer
Dim wRazao As String
Dim wEndereco As String
Dim WCOD_BARRA As String
Close #1
Open "LPT1" For Output As #1
        
WCOD_BARRA = CStr(numeroOS)
wRazao = "ELETRONICA DORVA LTDA"
wEndereco = "Rua 13 de Maio,122A-Sao Luiz"
WFONE = "Fone: 3268-7095 - 98113-1813"
WPRECO = ConverteValor(TOTAL_GERAL)
            
Print #1, "I8,A,032"
Print #1, "Q243,23" 'Q184 > 184 significa 184 dots, 1 mm = 8 Dots é a altura da etiqueta 184 Dots = 23 mm (2,3 cm) , 24 dots espaço entre etiquetas
Print #1, "q320"
Print #1, "rY"
Print #1, "S4"      'Determina a velocidade da impressão
Print #1, "D04"     'Determina o fator de escuridao da etiqueta
Print #1, "ZT"      'Determina a sequencia de impressão T = Top B = Button
Print #1, "JF"      'Disable Top Of Form Backup
Print #1, "OD"
Print #1, "R256,0"   'Determina a margem da impressora
Print #1, "WY"
Print #1, "N"       'Limpa a memoria da impressora a cada nova impressao
           
'A0 > COLUNA
'0  > LINHA
'0  > ROTAÇÃO
'3  > TIPO DE FONTE
'1  > MUTIPLICADOR ALTURA CARACTERES
'1  > MUTIPLICADOR LARGURA CARACTERES
'N  >
            
Print #1, "A25,5,0,3,1,2,N," & Chr(34) & wRazao & Chr(34)
Print #1, "A30,45,0,1,1,1,N," & Chr(34) & wEndereco & Chr(34)
Print #1, "A30,65,0,1,1,2,N," & Chr(34) & WFONE & Chr(34)
Print #1, "A30,95,0,1,1,1,N," & Chr(34) & Nome_Cliente & Chr(34)
Print #1, "A50,125,0,1,2,2,R," & Chr(34) & WPRECO & Chr(34) 'Preço
Print #1, "B50,155,0,1,2,2,50,B," & Chr(34) & WCOD_BARRA & Chr(34)
Print #1, "P1" '& WQTDE
    
Close #1

End Sub


Private Function ConverteValor(ByVal Valor As Currency) As String
Dim Inteiros As Currency
Dim Decimais As Currency
 
Dim sI As String
Dim sd As String
 
    Inteiros = Int(Valor)
    Decimais = (Valor - Inteiros) * 100
 
    sI = Format$(Inteiros, "0")
    sd = Format$(Decimais, "00")
 
    sI = Replace(sI, "0", "X")
    sI = Replace(sI, "1", "A")
    sI = Replace(sI, "2", "U")
    sI = Replace(sI, "3", "R")
    sI = Replace(sI, "4", "E")
    sI = Replace(sI, "5", "S")
    sI = Replace(sI, "6", "P")
    sI = Replace(sI, "7", "I")
    sI = Replace(sI, "8", "N")
    sI = Replace(sI, "9", "T")
 
    sd = Replace(sd, "0", "X")
    sd = Replace(sd, "1", "A")
    sd = Replace(sd, "2", "U")
    sd = Replace(sd, "3", "R")
    sd = Replace(sd, "4", "E")
    sd = Replace(sd, "5", "S")
    sd = Replace(sd, "6", "P")
    sd = Replace(sd, "7", "I")
    sd = Replace(sd, "8", "N")
    sd = Replace(sd, "9", "T")
 
    ConverteValor = sI & "," & sd
    
End Function
