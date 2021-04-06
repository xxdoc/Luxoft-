VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "1 5 3 4 2"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Text            =   "2"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Text            =   "5"
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Test Pairs Task"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Insertar array"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Insertar numero a comparar"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Insertar N"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As New Proyecto1.Class1

Private Sub Command1_Click()

obj.Pairs

End Sub

