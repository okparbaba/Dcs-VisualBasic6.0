VERSION 5.00
Begin VB.Form frmChanging 
   Caption         =   "Changing Caption to Command Button"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "Changing"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmChanging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    cmdChange.Caption = Text1.Text
End Sub
