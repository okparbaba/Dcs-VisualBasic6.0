VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "C&hange Case"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CheckBox chkLower 
      Caption         =   "Lower Case"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkUpper 
      Caption         =   "Upper Case"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a sentance :"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLower_Click()
    chkLower.Value = 0
End Sub

Private Sub chkUpper_Click()
    chkUpper.Value = 0
End Sub

Private Sub cmdChange_Click()
    If chkUpper.Value = 1 Then
        Text1 = Format(Text1.Text, ">")
    ElseIf chkLower.Value = 1 Then
        Text1 = Format(Text1.Text, "<")
    Else
        MsgBox "Please, Choose one case to change!", vbCritical + vbOKCancel, "Message Box"
    End If
    
End Sub

Private Sub cmdClear_Click()
    Text1.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub
