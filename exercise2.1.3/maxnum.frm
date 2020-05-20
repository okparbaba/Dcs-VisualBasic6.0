VERSION 5.00
Begin VB.Form frmMaxnum 
   Caption         =   "Maximum Number"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtResult 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtNo3 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtNo2 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtNo1 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Result"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Number3"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Number2"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Number1"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmMaxnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
    Dim x1, x2, x3, xmax As Double
    x1 = Val(txtNo1.Text)
    x2 = Val(txtNo2.Text)
    x3 = Val(txtNo3.Text)
    
    xmax = x1
    If x2 > xmax Then
    xmax = x2
    End If
    If x3 > xmax Then
    xmax = x3
    End If
    txtResult.Text = CStr(xmax)
    
End Sub
