VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblResult 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "="
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtSNum 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cboOperator 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtFNum 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
    Dim tmp As Double
    If cboOperator.ListIndex = -1 Then
        MsgBox "Operator?", vbQuestion + vbOKOnly
        cboOperator.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtFNum.Text) = False Then
        MsgBox "Invalid Number.", vbExclamation + vbOKOnly
        txtFNum.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtSNum.Text) = False Then
        MsgBox "Invalid Number.", vbExclamation + vbOKOnly
        txtSNum.SetFocus
        Exit Sub
    End If

    tmp = CLng(txtSNum.Text)
    If cboOperator.Text = "/" And tmp = 0 Then
        MsgBox "Invalid Divisor.", vbCritical + vbOKOnly
        txtSNum.SetFocus
        Exit Sub
    End If
    Select Case cboOperator.Text
        Case "+"
            lblResult = CStr(CLng(txtFNum) + CLng(txtSNum))
        Case "-"
            lblResult = CStr(CLng(txtFNum) - CLng(txtSNum))
        Case "*"
            lblResult = CStr(CLng(txtFNum) * CLng(txtSNum))
        Case "/"
            lblResult = CStr(CLng(txtFNum) / CLng(txtSNum))
    End Select
End Sub

Private Sub Form_Load()
    cboOperator.AddItem "+"
    cboOperator.AddItem "-"
    cboOperator.AddItem "*"
    cboOperator.AddItem "/"
End Sub
