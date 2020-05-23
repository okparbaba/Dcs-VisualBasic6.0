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
   Begin VB.ListBox lstCollection 
      Height          =   2205
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtToPut 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdAdd_Click()
    lstCollection.AddItem (txtToPut.Text)
    txtToPut.Text = ""
    txtToPut.SetFocus
End Sub

Private Sub cmdClear_Click()
    lstCollection.Clear
End Sub

Private Sub cmdDelete_Click()
    lstCollection.RemoveItem (i)
End Sub
