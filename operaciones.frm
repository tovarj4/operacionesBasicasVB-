VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Operaciones basicas"
   ClientHeight    =   10785
   ClientLeft      =   7260
   ClientTop       =   915
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   9495
   Begin VB.ListBox ListTablas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   2400
      TabIndex        =   16
      Top             =   5760
      Width           =   3855
   End
   Begin VB.ComboBox ComTablas 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H008080FF&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      MaskColor       =   &H000000FF&
      TabIndex        =   9
      Top             =   10080
      Width           =   1815
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "Divide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdMultiplica 
      Caption         =   "Multiplica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdResta 
      Caption         =   "Resta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdSuma 
      Caption         =   "Suma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtNum2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNum1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblTablas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tablas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   3720
      TabIndex        =   14
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Operaciones Basicas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label Label3 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Casilla DOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Casilla UNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblResultado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblSigno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaramos las variables uno y dos, que son con las que se realizan las operaciones
Dim uno As Double, dos As Double
'Funcion que validan que los Text Box no esten vacios
Function r() As Boolean
' Si es vacio envia mensaje de error y devuelve false
If txtNum1.Text = "" Then
MsgBox "Debes ingresar un número en la casilla UNO ", 48
r = False

    ElseIf txtNum2.Text = "" Then
    MsgBox "Debes ingresar un número en la casilla DOS ", 48
    Else
    'si ambos TextBox estan llenos asigna el valor a las variables uno y dos, devuelve true
    uno = Val(txtNum1.Text)
    dos = Val(txtNum2.Text)
    r = True
    
End If
End Function
'esta funcion valida si el resultado es negativo lo pone en color rojo
Function validaNegativo(n As Double)
If n < 0 Then
lblResultado.ForeColor = &HFF&
Else
lblResultado.ForeColor = &H0&
End If
End Function
' Funcion Valida el tipo de operacion y se la asigna a la variable resultado
Function resultado(ope As String) As Double
    
    Select Case ope
    Case "+" 'Suma
    resultado = uno + dos
    
    Case "-" 'Resta
    resultado = uno - dos
    
    Case "/" 'Divide
    resultado = uno / dos
    
    Case "*" 'Multiplica
    resultado = uno * dos
    
    End Select
    'Valida que el resultado no sea negativo
validaNegativo (resultado)
End Function



Private Sub ComTablas_Click()
Dim i As Integer, c As Integer, r As Integer
    i = ComTablas.ListIndex
      ListTablas.Clear
    For c = 1 To 100
        r = i * c
      
        
        ListTablas.AddItem (CStr(i) & "  X  " & CStr(c) & "  =  " & CStr(r))
    Next
    
End Sub

Private Sub Form_Load()
Dim i  As Integer
For i = 0 To 100
    ComTablas.AddItem CStr(i)
Next i
End Sub



'Validan en Keypress Que el Textbox solo acepte numeros y punto
Private Sub txtNum1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Or (KeyAscii >= 33) And (KeyAscii <= 45) Or (KeyAscii >= 58) And (KeyAscii <= 100) Or _
(KeyAscii >= 91) And (KeyAscii <= 96) Or (KeyAscii >= 123) And (KeyAscii <= 126) Then
MsgBox "Solo Acepta numeros"
KeyAscii = 8
End If
End Sub
'Validan en Keypress Que el Textbox solo acepte numeros y punto
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Or (KeyAscii >= 33) And (KeyAscii <= 45) Or (KeyAscii >= 58) And (KeyAscii <= 100) Or _
(KeyAscii >= 91) And (KeyAscii <= 96) Or (KeyAscii >= 123) And (KeyAscii <= 126) Then
MsgBox "Solo Acepta numeros"
KeyAscii = 8
End If
End Sub
Private Sub cmdSuma_Click()
If r = True Then
    lblSigno = "+"
    lblResultado.Caption = resultado("+")
Else
End If
 
End Sub
Private Sub cmdResta_Click()
If r = True Then
    lblSigno = "-"
    lblResultado.Caption = resultado("-")
Else
End If
 
End Sub
Private Sub cmdDivide_Click()
If r = True Then
    lblSigno = "/"
    lblResultado.Caption = resultado("/")
Else
End If
 
End Sub
Private Sub cmdMultiplica_Click()
If r = True Then
    lblSigno = "*"
    lblResultado.Caption = resultado("*")
Else
End If
 
End Sub

Private Sub cmdCerrar_Click()

        Unload Form1

End Sub
'Fin
