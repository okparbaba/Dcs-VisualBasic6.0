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
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    Dim a, b As Double
    Dim i, j As Integer
    For i = -6 To 6
    j = i + 1
    a = f(i)
    b = f(j)
        MsgBox ("x1 =" & CStr(i) & " y1 = " & CStr(a) & _
        vbNewLine & "x2 = " & CStr(j) & "y2 = " & CStr(b))
       If a * b < 0 Then
           MsgBox ("There is a solution between x1 = " & CStr(i) & ", x2= " & CStr(j) & _
           vbNewLine & "with y1 = " & CStr(a) & ",    y2 = " & CStr(b))
           txtResult.Text = "There is a solution in " + CStr(i) + "and" + CStr(j)
           Exit For     ' This stops the looping if a solution were found
        End If
        Next i
End Sub
