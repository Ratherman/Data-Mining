VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9915
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "執行"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   3960
      Width           =   5895
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   3240
      Width           =   6015
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   7815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "三角形面積 ="
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "三角形周長 ="
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "三角形判斷"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "鄭大嘉"
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "一年碩班"
      Height          =   735
      Left            =   4080
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "C ="
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "B ="
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "A ="
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim A, B, C As Single
Dim T As Boolean

A = Val(Text1.Text)
B = Val(Text2.Text)
C = Val(Text3.Text)

L = A + B + C
S = L / 2
D = S * (S - A) * (S - B) * (S - C)

If D <= 0 Then
    Text4.Text = "A, B, C 三邊不能構成一個三角形!"
    T = False
Else
    Text4.Text = "A, B, C 三邊可以構成一個三角形!"
    T = True
End If

If T = True Then
    F = Sqr(D)
    Text5.Text = Str(L)
    Text6.Text = Str(F)
Else
    If T = False Then
        Text5.Text = "不能構成三角形!"
        Text6.Text = "不能構成三角形!"
    End If

End If

End Sub

Private Sub Command2_Click()
End
End Sub
