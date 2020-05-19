VERSION 5.00
Begin VB.Form frmCircle 
   AutoRedraw      =   -1  'True
   Caption         =   "To Draw Circle By hsb"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hbsRadius 
      Height          =   495
      Left            =   1440
      Max             =   1250
      Min             =   1
      TabIndex        =   0
      Top             =   3720
      Value           =   500
      Width           =   6495
   End
   Begin VB.Label lblArea 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4560
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Area now = "
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
      Left            =   3120
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "frmCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI As Single = 3.14159
Dim Radius As Single

Private Sub hbsRadius_Change()
    Dim CY As Long
    Dim CX As Long
    
    CX = Me.ScaleWidth / 2
    CY = (Me.ScaleHeight - 600) / 2
    
    Radius = hbsRadius.Value
    Me.Cls
    Me.Circle (CX, CY), Radius
    lblArea.Caption = CStr(Radius * Radius * PI)
    
End Sub

Private Sub hbsRadius_Scroll()
    Call hbsRadius_Change
End Sub

Private Sub Form_Activite()
    Call hbsRadius_Change
End Sub

