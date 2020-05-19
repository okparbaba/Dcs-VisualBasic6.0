VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Using Frame"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Text            =   "hello"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Height          =   1695
         Left            =   3000
         Picture         =   "frmFrame.frx":0000
         ScaleHeight     =   641.176
         ScaleMode       =   0  'User
         ScaleWidth      =   457.426
         TabIndex        =   3
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Picture1.Picture = LoadPicture("C:\convo.jpg")
    Label1.Caption = Date
    Label2.Caption = Time
End Sub

