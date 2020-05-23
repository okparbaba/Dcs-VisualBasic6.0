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
   Begin VB.ComboBox cobUniy 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "Choose University"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtIntake 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cobUniy_Change()
    txtIntake.Text = cobUniy.Text
    
End Sub

Private Sub Form_Load()
    cobUniy.AddItem ("Bago University")
    cobUniy.AddItem ("Yangon University")
    cobUniy.AddItem ("Sittwe University")
End Sub
