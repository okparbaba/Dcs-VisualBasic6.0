VERSION 5.00
Begin VB.Form frmListBox 
   Caption         =   "Adding Text To Listbox"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddList 
      Caption         =   "Add List"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activiate()
    Text1.SetFocus
End Sub

Private Sub cmdAddList_Click()
    List1.AddItem Text1.Text
    Text1.SetFocus
    Text1.Text = ""
End Sub
