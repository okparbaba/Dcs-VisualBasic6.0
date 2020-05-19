VERSION 5.00
Begin VB.Form frmDateTime 
   Caption         =   "Date & Time Program"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblTime 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Time :"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    End
End Sub

Private Sub Form_Load()
    lblDate.Caption = Date
    lblTime.Caption = Time
End Sub
