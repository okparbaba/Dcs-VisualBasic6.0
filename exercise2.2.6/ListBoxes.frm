VERSION 5.00
Begin VB.Form frmListBox 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdBackAll 
      Caption         =   "<<"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveAll 
      Caption         =   ">>"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox lstLeft 
      Height          =   2400
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Dim i As Integer
    If lstRight.SelCount > 0 Then           'If .. Then .. End If
        i = lstRight.ListCount - 1

        Do While i >= 0                 'Do While .. .. Loop
            If lstRight.Selected(i) = True Then
                lstLeft.AddItem lstRight.List(i)
                lstRight.RemoveItem i
            End If
            i = i - 1
        Loop
    End If
End Sub

Private Sub cmdBackAll_Click()
    Dim i As Integer
    For i = 0 To lstRight.ListCount - 1
        lstLeft.AddItem lstRight.List(i)
    Next i
    
    Do Until lstRight.ListCount = 0
        lstRight.RemoveItem (0)
    Loop

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click()
    Dim i As Integer
    If lstLeft.SelCount > 0 Then
        i = lstLeft.ListCount - 1
        Do While i >= 0                 'Do While .. .. Loop
            If lstLeft.Selected(i) = True Then
                lstRight.AddItem lstLeft.List(i)
                lstLeft.RemoveItem i
            End If
            i = i - 1
        Loop
    End If

End Sub

Private Sub cmdMoveAll_Click()
    Dim i As Integer
    For i = 0 To lstLeft.ListCount - 1      'For .. Next
        lstRight.AddItem lstLeft.List(i)
    Next i
    
    Do Until lstLeft.ListCount = 0          'Do Until  … Loop
        lstLeft.RemoveItem (0)
    Loop

End Sub

Private Sub Form_Load()
    Dim i As Byte
    For i = 1 To 10
        lstLeft.AddItem "Line - " & CStr(i)
    Next i
End Sub
