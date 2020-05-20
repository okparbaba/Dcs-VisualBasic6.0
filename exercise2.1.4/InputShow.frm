VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtLn 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShow_Click()
    Dim x As String
    x = Trim(txtLn.Text)
    Select Case x
        Case "One"
        MsgBox "Your Input is 1"
        Case "Two"
        MsgBox "Your Input is 2"
        Case "Three"
        MsgBox "Your Input is 3"
        Case Else
        MsgBox "Your Input is 1 nor 2 nor 3"
    End Select
End Sub
