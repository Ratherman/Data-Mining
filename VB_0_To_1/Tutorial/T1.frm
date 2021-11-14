VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6870
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "VB程式練習"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "機械系二專部一年三班"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   3840
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()

PI = 3.14159
R = 50
C = 2 * PI * PI
A = PI * R ^ 2

Print "圓半徑 ="; R
Print "圓周長 ="; C
Print "圓面積 ="; A
End Sub
