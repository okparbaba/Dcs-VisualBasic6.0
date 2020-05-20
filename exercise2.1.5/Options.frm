VERSION 5.00
Begin VB.Form frmOption 
   Caption         =   "Option Program"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdShowMsgBox 
      Caption         =   "Show Message Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   3720
      Width           =   2655
   End
   Begin VB.OptionButton optYesNo 
      Caption         =   "Yes or No"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.OptionButton optOKCancel 
      Caption         =   "OK Cancel"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Button Group"
      Height          =   1935
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   2295
      Begin VB.OptionButton optOKOnly 
         Caption         =   "OK only"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.OptionButton optExclam 
      Caption         =   "Exclamation"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Case"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2175
      Begin VB.OptionButton optInfo 
         Caption         =   "Information"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "You have clicked Yes Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdShowMsgBox_Click()
    Dim K As Integer
    Dim IconType As Integer
    Dim ButtonGroup As Integer
    
    If optInfo.Value = True Then
        IconType = vbInformation
    ElseIf optExclam.Value = True Then
        IconType = vbExclamation
    End If
    
    If optOKOnly.Value = True Then
        ButtonGroup = vbOKOnly
    ElseIf optOKCancel.Value = True Then
        ButtonGroup = vbOKCancel
    ElseIf optYesNo.Value = True Then
        ButtonGroup = vbYesNo
    End If
    
    K = MsgBox("Welcome to VB 6.0 !", IconType + ButtonGroup, "Option Testing")

    Select Case K
        Case vbOK
            lblInfo.Caption = "You have clicked OK button."
        Case vbCancel
            lblInfo.Caption = "You have clicked Cancel button."
        Case vbYes
            lblInfo.Caption = "You have clicked Yes button."
        Case vbNo
            lblInfo.Caption = "You have clicked No button."
    End Select

End Sub
