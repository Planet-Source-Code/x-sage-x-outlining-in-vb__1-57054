VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Outlining is fun!"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Common 
      Left            =   720
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "A"
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Label2 
      Caption         =   "Middle Color:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Outline Color:"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Lettera 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Lettera 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   105
   End
   Begin VB.Label Lettera 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   135
      Width           =   105
   End
   Begin VB.Label Lettera 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Lettera 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lett As Integer

Private Sub Picture1_DblClick()
Common.ShowColor
Picture1.BackColor = Common.Color
Lettera(1).ForeColor = Picture1.BackColor
Lettera(2).ForeColor = Picture1.BackColor
Lettera(3).ForeColor = Picture1.BackColor
Lettera(4).ForeColor = Picture1.BackColor
End Sub

Private Sub Picture2_DblClick()
Common.ShowColor
Picture2.BackColor = Common.Color
Lettera(0).ForeColor = Picture2.BackColor
End Sub

Private Sub Text1_Change()
Lettera(0).Caption = Text1.Text
Lettera(1).Caption = Text1.Text
Lettera(2).Caption = Text1.Text
Lettera(3).Caption = Text1.Text
Lettera(4).Caption = Text1.Text
Lettera(0).ForeColor = Picture2.BackColor
Lettera(1).ForeColor = Picture1.BackColor
Lettera(2).ForeColor = Picture1.BackColor
Lettera(3).ForeColor = Picture1.BackColor
Lettera(4).ForeColor = Picture1.BackColor
End Sub
