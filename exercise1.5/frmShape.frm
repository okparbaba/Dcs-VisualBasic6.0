VERSION 5.00
Begin VB.Form frmShape 
   Caption         =   "Changing Shape"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeShape 
      Caption         =   "Change Shape"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   3000
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChangeShape_Click()
    Shape1.Shape = Val(Text1.Text)
End Sub
