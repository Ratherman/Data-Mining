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
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' Declare a single-dimension array of 5 numbers
Dim numbers(4) As Integer

' Declare a 6 x 6 multidimensional array
Dim matrix(5, 5) As Double

' Delcare an array with 7 elements.
Dim students(6) As Integer

' Assign values to each element.
students(0) = 23
students(1) = 19
students(2) = 21
students(3) = 17
students(4) = 19
students(5) = 20
students(6) = 22


' Display the value of each element
For ctr = 0 To 6
    List1.AddItem (Str(students(ctr)))
    
Next ctr


End Sub

Private Sub List1_Click()

End Sub
