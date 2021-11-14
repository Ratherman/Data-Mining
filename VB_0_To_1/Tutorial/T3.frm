VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   9750
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   615
      Left            =   3840
      TabIndex        =   16
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "四則運算"
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   5175
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   960
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "A / B ="
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "A * B ="
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "A - B ="
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "A + B ="
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "鄭中嘉 L46104020"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "地科系 碩一"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "B ="
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "A ="
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim A, B As Single

A = Val(Text1.Text)
B = Val(Text2.Text)

C1 = A + B
C2 = A - B
C3 = A * B
If B = 0 Then
    MsgBox "除法分母不可為零!", , "錯誤訊息"
Else
    C4 = A / B
End If

Text3.Text = Str(C1)
Text4.Text = Str(C2)
Text5.Text = Str(C3)

If B = 0 Then
    Text6.Text = "分母為0,錯誤!"
Else
    Text6.Text = Str(C4)
End If

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub
