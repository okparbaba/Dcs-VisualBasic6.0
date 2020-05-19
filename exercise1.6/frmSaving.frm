VERSION 5.00
Begin VB.Form frmSaving 
   Caption         =   "Saving"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1700
      Width           =   1095
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtInterest 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtDeposit 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Total"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1700
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Interest"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Deposite"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Deposit = Val(txtDeposit.Text)
    IR = Val(txtInterest.Text) / 100
    yr = Val(txtYear.Text)
    TotalAmount = Deposit * (1 + yr * IR)
    txtTotal.Text = Format(TotalAmount, "Fixed")
End Sub
