VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   9570
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   6360
      TabIndex        =   14
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "答案 S="
      Height          =   615
      Left            =   5040
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Text            =   "200"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "答案 S="
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Text            =   "99"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "答案 S="
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Text            =   "100"
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "N(偶數) ="
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      Caption         =   "計算 S = 2+4+6+8+...+N 的總和"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   9255
   End
   Begin VB.Label Label4 
      Caption         =   "N(奇數) = "
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "計算 S = 1+3+5+7+...+N的總和"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   9255
   End
   Begin VB.Label Label2 
      Caption         =   "N(整數)="
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "計算 S = 1+2+3+4+...+N 的總和"
      Height          =   180
      Left            =   3570
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

N = Val(Text1.Text)
S = 0
For I = 1 To N
    S = S + I
Next I
Text2.Text = Str(S)

End Sub

Private Sub Command2_Click()

N = Val(Text3.Text)
S = 0
For I = 1 To N Step 2
    S = S + I
Next I
Text4.Text = Str(S)

End Sub

Private Sub Command3_Click()

N = Val(Text5.Text)
S = 0
For I = 2 To N Step 2
    S = S + I
Next I
Text6.Text = Str(S)

End Sub

Private Sub Text1_Change()
Text2.Text = ""
End Sub

Private Sub Text3_Change()
Text4.Text = ""
End Sub

Private Sub Text5_Change()
Text6.Text = ""
End Sub
