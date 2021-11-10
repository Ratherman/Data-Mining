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
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "圓面積 ="
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "圓周長 ="
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "圓半徑 = "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_Change()

PI = 3.1416
R = Val(Text1.Text)

C = 2 * PI * R
A = PI * R ^ 2

Text2.Text = Str(C)
Text3.Text = Str(A)
End Sub
