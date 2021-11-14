VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   11760
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8220
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   16695
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "soybean-small.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Click:(1) Read (2) Forward Selection (3) Backward Selection"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Partition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Single
Dim j As Single
Dim k As Single

Dim in_file As String
Dim attributes(35, 46) As String
Dim single_entropy(35) As Double
Dim joint_entropy(35, 35) As Double
Dim symmetric_uncertainty(35, 35) As Double
Dim feat(35) As String


Private Sub Partition_click()

' 清空 List 1 的內容
List1.Clear

' (1) Read Txt File
List1.AddItem "Step 1: Read Txt File"

in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat(0), feat(1), feat(2), feat(3), feat(4), feat(5), feat(6), feat(7), feat(8), feat(9), feat(10), feat(11), feat(12), feat(13), feat(14), feat(15), feat(16), feat(17), feat(18), feat(19), feat(20), feat(21), feat(22), feat(23), feat(24), feat(25), feat(26), feat(27), feat(28), feat(29), feat(30), feat(31), feat(32), feat(33), feat(34), feat(35)
    
    For j = 0 To 35
        attributes(j, i) = feat(j)
    Next j
    i = i + 1
Loop
Close #1

' (2) Derive entropy for each column (attribute)
List1.AddItem "Step 2: Derive entropy for each column (attribute)"

For j = 0 To 35
    ' 取出特定一 column
    Dim col(46) As String
    For i = 0 To 46
        col(i) = attributes(j, i)
    Next
    
    ' 計數
    Dim count(10) As Integer
    For i = 0 To 10
        count(i) = 0
    Next
    
    For i = 0 To 46
        If col(i) = "0" Then
            count(0) = count(0) + 1
        ElseIf col(i) = "1" Then
            count(1) = count(1) + 1
        ElseIf col(i) = "2" Then
            count(2) = count(2) + 1
        ElseIf col(i) = "3" Then
            count(3) = count(3) + 1
        ElseIf col(i) = "4" Then
            count(4) = count(4) + 1
        ElseIf col(i) = "5" Then
            count(5) = count(5) + 1
        ElseIf col(i) = "6" Then
            count(6) = count(6) + 1
        ElseIf col(i) = "D1" Then
            count(7) = count(7) + 1
        ElseIf col(i) = "D2" Then
            count(8) = count(8) + 1
        ElseIf col(i) = "D3" Then
            count(9) = count(9) + 1
        ElseIf col(i) = "D4" Then
            count(10) = count(10) + 1
        End If
    Next
    
    '機率
    Dim prob_list(10) As Double
    For i = 0 To 10
        prob_list(i) = count(i) / 47
    Next
    
    'Entropy
    Dim entropy As Double
    Dim log_val As Double
    
    entropy = 0
    For i = 0 To 10
        If prob_list(i) = 0 Then
            entropy = entropy
        Else
            log_val = Log(prob_list(i)) / Log(2)
            entropy = entropy - prob_list(i) * log_val
        End If
    Next
    
    single_entropy(j) = entropy
Next

' (3) Derive entropy for each two different two columns (attributes)
List1.AddItem "Step 3: Derive joint entropy for each two different columns (attributes)"


For j = 0 To 35
For k = 0 To 35

' 取出特定兩個 column
Dim col_A(46) As String
Dim col_B(46) As String

For i = 0 To 46
    col_A(i) = attributes(j, i)
    col_B(i) = attributes(k, i)
Next

' 計數
Dim double_count(120) As Integer

For i = 0 To 120
    double_count(i) = 0
Next

For i = 0 To 46
    ' 0
    If col_A(i) = "0" And col_B(i) = "0" Then
        double_count(0) = double_count(0) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "1" Then
        double_count(1) = double_count(1) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "2" Then
        double_count(2) = double_count(2) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "3" Then
        double_count(3) = double_count(3) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "4" Then
        double_count(4) = double_count(4) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "5" Then
        double_count(5) = double_count(5) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "6" Then
        double_count(6) = double_count(6) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "D1" Then
        double_count(7) = double_count(7) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "D2" Then
        double_count(8) = double_count(8) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "D3" Then
        double_count(9) = double_count(9) + 1
    ElseIf col_A(i) = "0" And col_B(i) = "D4" Then
        double_count(10) = double_count(10) + 1
    ' 1
    ElseIf col_A(i) = "1" And col_B(i) = "0" Then
        double_count(11) = double_count(11) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "1" Then
        double_count(12) = double_count(12) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "2" Then
        double_count(13) = double_count(13) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "3" Then
        double_count(14) = double_count(14) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "4" Then
        double_count(15) = double_count(15) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "5" Then
        double_count(16) = double_count(16) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "6" Then
        double_count(17) = double_count(17) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "D1" Then
        double_count(18) = double_count(18) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "D2" Then
        double_count(19) = double_count(19) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "D3" Then
        double_count(20) = double_count(20) + 1
    ElseIf col_A(i) = "1" And col_B(i) = "D4" Then
        double_count(21) = double_count(21) + 1
    ' 2
    ElseIf col_A(i) = "2" And col_B(i) = "0" Then
        double_count(22) = double_count(22) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "1" Then
        double_count(23) = double_count(23) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "2" Then
        double_count(24) = double_count(24) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "3" Then
        double_count(25) = double_count(25) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "4" Then
        double_count(26) = double_count(26) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "5" Then
        double_count(27) = double_count(27) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "6" Then
        double_count(28) = double_count(28) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "D1" Then
        double_count(29) = double_count(29) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "D2" Then
        double_count(30) = double_count(30) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "D3" Then
        double_count(31) = double_count(31) + 1
    ElseIf col_A(i) = "2" And col_B(i) = "D4" Then
        double_count(32) = double_count(32) + 1
    ' 3
    ElseIf col_A(i) = "3" And col_B(i) = "0" Then
        double_count(33) = double_count(33) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "1" Then
        double_count(34) = double_count(34) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "2" Then
        double_count(35) = double_count(35) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "3" Then
        double_count(36) = double_count(36) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "4" Then
        double_count(37) = double_count(37) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "5" Then
        double_count(38) = double_count(38) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "6" Then
        double_count(39) = double_count(39) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "D1" Then
        double_count(40) = double_count(40) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "D2" Then
        double_count(41) = double_count(41) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "D3" Then
        double_count(42) = double_count(42) + 1
    ElseIf col_A(i) = "3" And col_B(i) = "D4" Then
        double_count(43) = double_count(43) + 1
    ' 4
    ElseIf col_A(i) = "4" And col_B(i) = "0" Then
        double_count(44) = double_count(44) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "1" Then
        double_count(45) = double_count(45) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "2" Then
        double_count(46) = double_count(46) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "3" Then
        double_count(47) = double_count(47) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "4" Then
        double_count(48) = double_count(48) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "5" Then
        double_count(49) = double_count(49) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "6" Then
        double_count(50) = double_count(50) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "D1" Then
        double_count(51) = double_count(51) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "D2" Then
        double_count(52) = double_count(52) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "D3" Then
        double_count(53) = double_count(53) + 1
    ElseIf col_A(i) = "4" And col_B(i) = "D4" Then
        double_count(54) = double_count(54) + 1
    ' 5
    ElseIf col_A(i) = "5" And col_B(i) = "0" Then
        double_count(55) = double_count(55) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "1" Then
        double_count(56) = double_count(56) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "2" Then
        double_count(57) = double_count(57) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "3" Then
        double_count(58) = double_count(58) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "4" Then
        double_count(59) = double_count(59) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "5" Then
        double_count(60) = double_count(60) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "6" Then
        double_count(61) = double_count(61) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "D1" Then
        double_count(62) = double_count(62) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "D2" Then
        double_count(63) = double_count(63) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "D3" Then
        double_count(64) = double_count(64) + 1
    ElseIf col_A(i) = "5" And col_B(i) = "D4" Then
        double_count(65) = double_count(65) + 1
    ' 6
    ElseIf col_A(i) = "6" And col_B(i) = "0" Then
        double_count(66) = double_count(66) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "1" Then
        double_count(67) = double_count(67) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "2" Then
        double_count(68) = double_count(68) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "3" Then
        double_count(69) = double_count(69) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "4" Then
        double_count(70) = double_count(70) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "5" Then
        double_count(71) = double_count(71) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "6" Then
        double_count(72) = double_count(72) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "D1" Then
        double_count(73) = double_count(73) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "D2" Then
        double_count(74) = double_count(74) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "D3" Then
        double_count(75) = double_count(75) + 1
    ElseIf col_A(i) = "6" And col_B(i) = "D4" Then
        double_count(76) = double_count(76) + 1
    ' D1
    ElseIf col_A(i) = "D1" And col_B(i) = "0" Then
        double_count(77) = double_count(77) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "1" Then
        double_count(78) = double_count(78) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "2" Then
        double_count(79) = double_count(79) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "3" Then
        double_count(80) = double_count(80) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "4" Then
        double_count(81) = double_count(81) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "5" Then
        double_count(82) = double_count(82) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "6" Then
        double_count(83) = double_count(83) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "D1" Then
        double_count(84) = double_count(84) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "D2" Then
        double_count(85) = double_count(85) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "D3" Then
        double_count(86) = double_count(86) + 1
    ElseIf col_A(i) = "D1" And col_B(i) = "D4" Then
        double_count(87) = double_count(87) + 1
    ' D2
    ElseIf col_A(i) = "D2" And col_B(i) = "0" Then
        double_count(88) = double_count(88) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "1" Then
        double_count(89) = double_count(89) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "2" Then
        double_count(90) = double_count(90) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "3" Then
        double_count(91) = double_count(91) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "4" Then
        double_count(92) = double_count(92) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "5" Then
        double_count(93) = double_count(93) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "6" Then
        double_count(94) = double_count(94) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "D1" Then
        double_count(95) = double_count(95) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "D2" Then
        double_count(96) = double_count(96) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "D3" Then
        double_count(97) = double_count(97) + 1
    ElseIf col_A(i) = "D2" And col_B(i) = "D4" Then
        double_count(98) = double_count(98) + 1
    ' D3
    ElseIf col_A(i) = "D3" And col_B(i) = "0" Then
        double_count(99) = double_count(99) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "1" Then
        double_count(100) = double_count(100) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "2" Then
        double_count(101) = double_count(101) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "3" Then
        double_count(102) = double_count(102) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "4" Then
        double_count(103) = double_count(103) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "5" Then
        double_count(104) = double_count(104) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "6" Then
        double_count(105) = double_count(105) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "D1" Then
        double_count(106) = double_count(106) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "D2" Then
        double_count(107) = double_count(107) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "D3" Then
        double_count(108) = double_count(108) + 1
    ElseIf col_A(i) = "D3" And col_B(i) = "D4" Then
        double_count(109) = double_count(109) + 1
    ' D4
    ElseIf col_A(i) = "D4" And col_B(i) = "0" Then
        double_count(110) = double_count(110) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "1" Then
        double_count(111) = double_count(111) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "2" Then
        double_count(112) = double_count(112) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "3" Then
        double_count(113) = double_count(113) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "4" Then
        double_count(114) = double_count(114) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "5" Then
        double_count(115) = double_count(115) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "6" Then
        double_count(116) = double_count(116) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "D1" Then
        double_count(117) = double_count(117) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "D2" Then
        double_count(118) = double_count(118) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "D3" Then
        double_count(119) = double_count(119) + 1
    ElseIf col_A(i) = "D4" And col_B(i) = "D4" Then
        double_count(120) = double_count(120) + 1
    End If
Next

'機率
Dim double_prob_list(120) As Double

For i = 0 To 120
    double_prob_list(i) = double_count(i) / 47
Next

'Joint Entropy

entropy = 0
For i = 0 To 120
    If double_prob_list(i) = 0 Then
        entropy = entropy
    Else
        log_val = Log(double_prob_list(i)) / Log(2)
        entropy = entropy - double_prob_list(i) * log_val
    End If
Next

joint_entropy(j, k) = entropy
Next
Next

' (4) Derive Symmetric Uncertainty for Each two different columns (Attributes)
List1.AddItem "Step 4: Derive Symmetric Uncertainty for Each two different columns (Attributes)"

Dim su As Double

For i = 0 To 35
For j = 0 To 35
    
    If i = j Then
        su = 1
    ElseIf single_entropy(i) + single_entropy(j) = 0 Then
        su = 0
    Else
        su = 2 * (single_entropy(i) + single_entropy(j) - joint_entropy(i, j)) / (single_entropy(i) + single_entropy(j))
    End If
        
    symmetric_uncertainty(i, j) = su
Next
Next


'(6) Forward Selection
List1.AddItem "Step 6: Forward Selection"

Dim must_chosen_subset(34) As Boolean
Dim possible_subset(34) As Boolean
Dim best_goodness As Double
Dim goodness_list(34) As Double
Dim exp_attr_list(34) As Integer
Dim count_same_vars As Integer
Dim chosen_subset(34) As Double
Dim good As Double
Dim statement As String
Dim idx As Integer
Dim current_best_attr As Integer
Dim current_bad_attr As Integer
Dim current_best_goodness As Double

For i = 0 To 34
    possible_subset(i) = True
    must_chosen_subset(i) = False
Next
best_goodness = 0
idx = -1

Do
    
    ' 一個 goodness 會對應到 一個 experiment 的 col，在開始前先行清空
    For i = 0 To 34
        goodness_list(i) = -1
        exp_attr_list(i) = -1
    Next

    ' 進行同樣 attr 數目的搜索
    count_same_vars = 0
    
    For j = 0 To 34
        
        If possible_subset(j) = True Then
        
            ' 每次搜索過相同 level 的 subset 時，需要清空 chosen_subset
            ' 在這邊一併加入實驗的 attr num
            For i = 0 To 34
            
                If i = j Then
                    chosen_subset(i) = i
                    exp_attr_list(count_same_vars) = i
                Else
                    chosen_subset(i) = -1 ' zero subset 的概念
                End If
            Next
            
            '把那些 最棒的 attr 加進來
            For i = 0 To 34
                If must_chosen_subset(i) = True Then
                    chosen_subset(i) = i
                End If
            Next
            
            good = get_goodness(chosen_subset)
            goodness_list(count_same_vars) = good
            count_same_vars = count_same_vars + 1
            
            statement = ""
            For i = 0 To 34
                If chosen_subset(i) <> -1 Then
                    statement = statement + Str(i) + " "
                End If
            Next
            List1.AddItem "Chosen Subset: [" & "" & statement & " " & "], Goodness:" & " " & good
        End If
    Next
    
    ' 找到 goodness　最大的 idx，並透過 exp_attr_list 找到 attr_num 和 goodness
    idx = argmax(goodness_list)
    current_best_attr = exp_attr_list(idx)
    current_best_goodness = goodness_list(idx)
    
    '設定停止條件
    ' (1) current_best_goodness 比 best_goodness 低
    If current_best_goodness < best_goodness Then
        Exit Do
    Else
        must_chosen_subset(current_best_attr) = True
        possible_subset(current_best_attr) = False
        best_goodness = current_best_goodness
    End If
Loop

statement = ""
For i = 0 To 34
    If must_chosen_subset(i) <> False Then
        statement = statement + Str(i) + " "
    End If
Next

List1.AddItem ""
List1.AddItem "Forward Selection Finish! Best Subset: [" & "" & statement & " " & "], Goodness:" & " " & best_goodness

'(7) Backward Selection
List1.AddItem "Step 7: Backward Selection"

Dim must_delete_subset(34) As Boolean

For i = 0 To 34
    possible_subset(i) = True
    must_delete_subset(i) = False
Next
best_goodness = 0
idx = -1

Do
    
    ' 一個 goodness 會對應到 一個 experiment 的 col，在開始前先行清空
    For i = 0 To 34
        goodness_list(i) = -1
        exp_attr_list(i) = -1
    Next

    ' 進行同樣 attr 數目的搜索
    count_same_vars = 0
    
    For j = 0 To 34
        
        If possible_subset(j) = True Then
        
            ' 每次搜索過相同 level 的 subset 時，需要填滿 chosen_subset
            ' 在這邊一併加入實驗的 attr num
            For i = 0 To 34
            
                If i <> j Then
                    chosen_subset(i) = i
                Else
                    exp_attr_list(count_same_vars) = i
                    chosen_subset(i) = -1 ' zero subset 的概念
                End If
            Next
            
            '把那些 要拿掉的 attr 拿掉
            For i = 0 To 34
                If must_delete_subset(i) = True Then
                    chosen_subset(i) = -1
                End If
            Next
            
            good = get_goodness(chosen_subset)
            goodness_list(count_same_vars) = good
            count_same_vars = count_same_vars + 1
            
            statement = ""
            For i = 0 To 34
                If chosen_subset(i) <> -1 Then
                    statement = statement + Str(i) + " "
                End If
            Next
            List1.AddItem "Chosen Subset: [" & "" & statement & " " & "], Goodness:" & " " & good
        End If
    Next
    
    ' 找到 goodness　最大的 idx，並透過 exp_attr_list 找到 attr_num 和 goodness
    idx = argmax(goodness_list)
    current_bad_attr = exp_attr_list(idx)
    current_best_goodness = goodness_list(idx)
    
    '設定停止條件
    ' (1) current_best_goodness 比 best_goodness 低
    If current_best_goodness < best_goodness Then
        Exit Do
    Else
        must_delete_subset(current_bad_attr) = True
        possible_subset(current_bad_attr) = False
        best_goodness = current_best_goodness
    End If
    
Loop

statement = ""
For i = 0 To 34
    If possible_subset(i) = True Then
        statement = statement + Str(i) + " "
    End If
Next

List1.AddItem ""
List1.AddItem "Backward Selection Finish! Best Subset: [" & "" & statement & " " & "], Goodness:" & " " & best_goodness



End Sub

' (5.5) Define argmax
Private Function argmax(arr)
    
    Dim max_num As Double
    Dim max_num_idx As Integer
    
    max_num = -1
    max_num_idx = -1
    
    For i = 0 To 34
        If arr(i) > max_num Then
            max_num = arr(i)
            max_num_idx = i
        End If
    Next
    
    argmax = max_num_idx
    
End Function

' (5) Define Goodness of an attribute subset S for classification C
Private Function get_goodness(chosen_subset)

Dim relevant_score As Double
relevant_score = 0

Dim attr_num As Variant
For Each attr_num In chosen_subset
    If attr_num = -1 Then
        relevant_score = relevant_score
    Else
        relevant_score = relevant_score + symmetric_uncertainty(attr_num, 35)
    End If
Next

Dim duplicate_score As Double
duplicate_score = 0

Dim attr_num_A As Variant
Dim attr_num_B As Variant
For Each attr_num_A In chosen_subset
    For Each attr_num_B In chosen_subset
        If attr_num_A = -1 Or attr_num_B = -1 Then
            duplicate_score = duplicate_score
        Else
            duplicate_score = duplicate_score + symmetric_uncertainty(attr_num_A, attr_num_B)
        End If
    Next
Next
duplicate_score = Sqr(duplicate_score)

get_goodness = relevant_score / duplicate_score
End Function

