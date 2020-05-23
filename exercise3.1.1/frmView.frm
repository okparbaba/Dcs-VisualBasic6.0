VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "Image Viewer"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Image1 
      Height          =   1815
      Left            =   1320
      ScaleHeight     =   1755
      ScaleWidth      =   3435
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   3720
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdShow_Click()
    Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
