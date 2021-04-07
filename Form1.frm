VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1080
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   3135
      TabIndex        =   8
      Top             =   -120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calcular"
      Height          =   495
      Left            =   1920
      MaskColor       =   &H8000000D&
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3360
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Task Test Pairs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insertar Array"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insertar numero a comparar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insertar el tamaño del Array"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As New Proyecto1.Class1
Dim numarray As String
Dim numcomparative As String
Dim arrays As String
Dim FutureVal As Variant
Private Sub Command1_Click()

'Assignment the values from the TexBox
numarray = Text1.Text
numcomparative = Text2.Text
arrays = Text3.Text

'calling the function
FutureVal = obj.Pairs(numarray, numcomparative, arrays)

'Clean all the Texboxs after the execute the function
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Validate than the user doesn't input letters
If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      MsgBox ("Este campo no se aceptan letras")
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'Validate than the user doesn't input letters
If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      MsgBox ("Este campo no se aceptan letras")
End If
End Sub

