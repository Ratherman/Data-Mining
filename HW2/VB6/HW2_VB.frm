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
   Begin VB.CommandButton Command1 
      Caption         =   "2. Perform Equal Frequency Discretization"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   5895
   End
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
      Width           =   10695
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
      Text            =   "Breast.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "1. Perform Equal Width Discretization"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   5895
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
Dim ee As Integer

Dim in_file As String
Dim attributes(9, 105) As Variant
Dim single_entropy(9) As Double
Dim joint_entropy(9, 9) As Double
Dim symmetric_uncertainty(9, 9) As Double

Dim feat_y As String
Dim feat_x(8) As Double
Dim attributes_EFD(9, 105) As Variant
Dim split_point_1 As Double
Dim split_point_2 As Double
Dim split_point_3 As Double
Dim split_point_4 As Double
Dim split_point_5 As Double
Dim split_point_6 As Double
Dim split_point_7 As Double
Dim split_point_8 As Double
Dim split_point_9 As Double
Dim num As Double

Dim str_list() As Variant
Dim num_list() As Variant

Dim elseif_num As Integer
Dim elseif_str As String
Dim elseif_num_list() As Variant
Dim elseif_str_list() As Variant




Private Sub Command1_Click()
' 清空 List 1 的內容
List1.Clear

' Read Txt File
List1.AddItem "Read Txt File"

in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat_x(0), feat_x(1), feat_x(2), feat_x(3), feat_x(4), feat_x(5), feat_x(6), feat_x(7), feat_x(8), feat_y
    
    For j = 0 To 8
        If j = 1 Or j = 2 Then
            attributes(j, i) = feat_x(j) * 100
        Else
            attributes(j, i) = feat_x(j)
        End If
    Next j
    attributes(9, i) = feat_y
    ' 全部列印出來看看
    ' List1.AddItem Str(attributes(0, i)) & "" & Str(attributes(1, i)) & "" & Str(attributes(2, i)) & "" & Str(attributes(3, i)) & "" & Str(attributes(4, i)) & "" & Str(attributes(5, i)) & "" & Str(attributes(6, i)) & "" & Str(attributes(7, i)) & "" & Str(attributes(8, i)) & "" & attributes(9, i)
    i = i + 1
    
Loop
Close #1

' (EFD 01) Equal-Frequency Discretization - Part I
List1.AddItem "(EFD 01) Equal-Frequency Discretization"
Dim attributes_sort(9, 105) As Variant
Dim column(105) As Double
Dim column_sort As Variant
Dim specific_col_sort(105) As Variant
Dim specific_col(105) As Variant


For i = 0 To 9
    If i <> 9 Then
        
        For j = 0 To 105
            column(j) = attributes(i, j)
        Next
        
        column_sort = sort_col(column)
        
        For j = 0 To 105
            attributes_sort(i, j) = column_sort(j)
        Next
        
    Else
        
        For j = 0 To 105
            attributes_sort(9, j) = attributes(9, j)
        Next
    End If
Next

' (EFD 02) Equal-Frequency Discretization - Part II
Dim discrete_attr(105) As Integer

For j = 0 To 9

    For i = 0 To 105
        specific_col_sort(i) = attributes_sort(j, i)
        specific_col(i) = attributes(j, i)
    Next
    
    If j <> 9 Then
        split_point_1 = specific_col_sort(1 * 10)
        split_point_2 = specific_col_sort(2 * 10)
        split_point_3 = specific_col_sort(3 * 10)
        split_point_4 = specific_col_sort(4 * 10)
        split_point_5 = specific_col_sort(5 * 10)
        split_point_6 = specific_col_sort(6 * 10)
        split_point_7 = specific_col_sort(7 * 10)
        split_point_8 = specific_col_sort(8 * 10)
        split_point_9 = specific_col_sort(9 * 10)
    
        ' transfer attributes into proper interval
        For i = 0 To 105
            discrete_attr(i) = -100
        Next
        
        For i = 0 To 105
            num = specific_col(i)
            
            If num >= split_point_1 Then
                discrete_attr(i) = 0
            ElseIf num >= split_point_2 And num < split_point_1 Then
                discrete_attr(i) = 1
            ElseIf num >= split_point_3 And num < split_point_2 Then
                discrete_attr(i) = 2
            ElseIf num >= split_point_4 And num < split_point_3 Then
                discrete_attr(i) = 3
            ElseIf num >= split_point_5 And num < split_point_4 Then
                discrete_attr(i) = 4
            ElseIf num >= split_point_6 And num < split_point_5 Then
                discrete_attr(i) = 5
            ElseIf num >= split_point_7 And num < split_point_6 Then
                discrete_attr(i) = 6
            ElseIf num >= split_point_8 And num < split_point_7 Then
                discrete_attr(i) = 7
            ElseIf num >= split_point_9 And num < split_point_8 Then
                discrete_attr(i) = 8
            ElseIf num < split_point_9 Then
                discrete_attr(i) = 9
            End If
        Next
        
        For i = 0 To 105
            attributes_EFD(j, i) = discrete_attr(i)
        Next
    Else
        For i = 0 To 105
            attributes_EFD(j, i) = attributes(j, i)
        Next
    End If
Next

' (2) Derive entropy for each column (attribute)
List1.AddItem "(EFD 02) Derive entropy for each column (attribute)"



str_list() = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "car", "fad", "mas", "gla", "con", "adi")
num_list() = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)

For j = 0 To 9
    ' 取出特定一 column
    Dim col(105) As String
    For i = 0 To 105
        col(i) = attributes_EFD(j, i)
    Next
    
    ' 計數
    Dim count(15) As Integer
    For i = 0 To 15
        count(i) = 0
    Next
    
    For i = 0 To 105
        For k = 0 To 15
            If col(i) = str_list(k) Then
                count(num_list(k)) = count(num_list(k)) + 1
            End If
        Next
    Next
    
    '機率
    Dim prob_list(15) As Double
    For i = 0 To 15
        prob_list(i) = count(i) / 106
    Next
    
    'Entropy
    Dim entropy As Double
    Dim log_val As Double
    
    entropy = 0
    For i = 0 To 15
        If prob_list(i) = 0 Then
            entropy = entropy
        Else
            log_val = Log(prob_list(i)) / Log(2)
            entropy = entropy - prob_list(i) * log_val
        End If
    Next
    
    single_entropy(j) = entropy
    
Next


' (EWD 03) Derive entropy for each two different two columns (attributes)
List1.AddItem "(EFD 03) Derive joint entropy for each two different columns (attributes)"


For j = 0 To 9
For k = 0 To 9

' 取出特定兩個 column
Dim col_A(105) As String
Dim col_B(105) As String

For i = 0 To 105
    col_A(i) = attributes_EFD(j, i)
    col_B(i) = attributes_EFD(k, i)
Next

' 計數
Dim double_count(255) As Integer

For i = 0 To 255
    double_count(i) = 0
Next



elseif_num_list() = Array(10, 11, 12, 13, 14, 15)
elseif_str_list() = Array("car", "fad", "mas", "gla", "con", "adi")

For i = 0 To 105
    For ee = 0 To 9
        ' ee 表示 0 ~ 9
        elseif_num = ee
        If Val(col_A(i)) = elseif_num And col_B(i) = "0" Then
            double_count(16 * elseif_num + 0) = double_count(16 * elseif_num + 0) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "1" Then
            double_count(16 * elseif_num + 1) = double_count(16 * elseif_num + 1) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "2" Then
            double_count(16 * elseif_num + 2) = double_count(16 * elseif_num + 2) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "3" Then
            double_count(16 * elseif_num + 3) = double_count(16 * elseif_num + 3) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "4" Then
            double_count(16 * elseif_num + 4) = double_count(16 * elseif_num + 4) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "5" Then
            double_count(16 * elseif_num + 5) = double_count(16 * elseif_num + 5) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "6" Then
            double_count(16 * elseif_num + 6) = double_count(16 * elseif_num + 6) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "7" Then
            double_count(16 * elseif_num + 7) = double_count(16 * elseif_num + 7) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "8" Then
            double_count(16 * elseif_num + 8) = double_count(16 * elseif_num + 8) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "9" Then
            double_count(16 * elseif_num + 9) = double_count(16 * elseif_num + 9) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "car" Then
            double_count(16 * elseif_num + 10) = double_count(16 * elseif_num + 10) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "fad" Then
            double_count(16 * elseif_num + 11) = double_count(16 * elseif_num + 11) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "mas" Then
            double_count(16 * elseif_num + 12) = double_count(16 * elseif_num + 12) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "gla" Then
            double_count(16 * elseif_num + 13) = double_count(16 * elseif_num + 13) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "con" Then
            double_count(16 * elseif_num + 14) = double_count(16 * elseif_num + 14) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "adi" Then
            double_count(16 * elseif_num + 15) = double_count(16 * elseif_num + 15) + 1
        End If
    
    Next
    
    
    For ee = 0 To 5
        ' 10 ~ 15
        elseif_num = elseif_num_list(ee)
        elseif_str = elseif_str_list(ee)
        If col_A(i) = elseif_str And col_B(i) = "0" Then
            double_count(16 * elseif_num + 0) = double_count(16 * elseif_num + 0) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "1" Then
            double_count(16 * elseif_num + 1) = double_count(16 * elseif_num + 1) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "2" Then
            double_count(16 * elseif_num + 2) = double_count(16 * elseif_num + 2) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "3" Then
            double_count(16 * elseif_num + 3) = double_count(16 * elseif_num + 3) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "4" Then
            double_count(16 * elseif_num + 4) = double_count(16 * elseif_num + 4) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "5" Then
            double_count(16 * elseif_num + 5) = double_count(16 * elseif_num + 5) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "6" Then
            double_count(16 * elseif_num + 6) = double_count(16 * elseif_num + 6) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "7" Then
            double_count(16 * elseif_num + 7) = double_count(16 * elseif_num + 7) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "8" Then
            double_count(16 * elseif_num + 8) = double_count(16 * elseif_num + 8) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "9" Then
            double_count(16 * elseif_num + 9) = double_count(16 * elseif_num + 9) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "car" Then
            double_count(16 * elseif_num + 10) = double_count(16 * elseif_num + 10) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "fad" Then
            double_count(16 * elseif_num + 11) = double_count(16 * elseif_num + 11) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "mas" Then
            double_count(16 * elseif_num + 12) = double_count(16 * elseif_num + 12) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "gla" Then
            double_count(16 * elseif_num + 13) = double_count(16 * elseif_num + 13) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "con" Then
            double_count(16 * elseif_num + 14) = double_count(16 * elseif_num + 14) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "adi" Then
            double_count(16 * elseif_num + 15) = double_count(16 * elseif_num + 15) + 1
        End If
    Next
Next

'機率
Dim double_prob_list(255) As Double

For i = 0 To 255
    double_prob_list(i) = double_count(i) / 106
Next

'Joint Entropy

entropy = 0
For i = 0 To 255
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

' (EWD 04) Derive Symmetric Uncertainty for Each two different columns (Attributes)
List1.AddItem "(EFD 04) Derive Symmetric Uncertainty for Each two different columns (Attributes)"

Dim su As Double

For i = 0 To 9
For j = 0 To 9
    
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


'(EWD 05) Forward Selection
List1.AddItem "(EFD 05) Forward Selection"

Dim must_chosen_subset(8) As Boolean
Dim possible_subset(8) As Boolean
Dim best_goodness As Double
Dim goodness_list(8) As Double
Dim exp_attr_list(8) As Integer
Dim count_same_vars As Integer
Dim chosen_subset(8) As Double
Dim good As Double
Dim statement As String
Dim idx As Integer
Dim current_best_attr As Integer
Dim current_bad_attr As Integer
Dim current_best_goodness As Double

For i = 0 To 8
    possible_subset(i) = True
    must_chosen_subset(i) = False
Next
best_goodness = 0
idx = -1

Do
    
    ' 一個 goodness 會對應到 一個 experiment 的 col，在開始前先行清空
    For i = 0 To 8
        goodness_list(i) = -1
        exp_attr_list(i) = -1
    Next

    ' 進行同樣 attr 數目的搜索
    count_same_vars = 0
    
    For j = 0 To 8
        
        If possible_subset(j) = True Then
        
            ' 每次搜索過相同 level 的 subset 時，需要清空 chosen_subset
            ' 在這邊一併加入實驗的 attr num
            For i = 0 To 8
            
                If i = j Then
                    chosen_subset(i) = i
                    exp_attr_list(count_same_vars) = i
                Else
                    chosen_subset(i) = -1 ' zero subset 的概念
                End If
            Next
            
            '把那些 最棒的 attr 加進來
            For i = 0 To 8
                If must_chosen_subset(i) = True Then
                    chosen_subset(i) = i
                End If
            Next
            
            good = get_goodness(chosen_subset)
            goodness_list(count_same_vars) = good
            count_same_vars = count_same_vars + 1
            
            statement = ""
            For i = 0 To 8
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
For i = 0 To 8
    If must_chosen_subset(i) <> False Then
        statement = statement + Str(i) + " "
    End If
Next

List1.AddItem ""
List1.AddItem "Forward Selection Finish! Best Subset: [" & "" & statement & " " & "], Goodness:" & " " & best_goodness


'(EFD 06) Backward Selection
List1.AddItem "(EFD 06) Backward Selection"

Dim must_delete_subset(8) As Boolean

For i = 0 To 8
    possible_subset(i) = True
    must_delete_subset(i) = False
Next
best_goodness = 0
idx = -1

Do
    
    ' 一個 goodness 會對應到 一個 experiment 的 col，在開始前先行清空
    For i = 0 To 8
        goodness_list(i) = -1
        exp_attr_list(i) = -1
    Next

    ' 進行同樣 attr 數目的搜索
    count_same_vars = 0
    
    For j = 0 To 8
        
        If possible_subset(j) = True Then
        
            ' 每次搜索過相同 level 的 subset 時，需要填滿 chosen_subset
            ' 在這邊一併加入實驗的 attr num
            For i = 0 To 8
            
                If i <> j Then
                    chosen_subset(i) = i
                Else
                    exp_attr_list(count_same_vars) = i
                    chosen_subset(i) = -1 ' zero subset 的概念
                End If
            Next
            
            '把那些 要拿掉的 attr 拿掉
            For i = 0 To 8
                If must_delete_subset(i) = True Then
                    chosen_subset(i) = -1
                End If
            Next
            
            good = get_goodness(chosen_subset)
            goodness_list(count_same_vars) = good
            count_same_vars = count_same_vars + 1
            
            statement = ""
            For i = 0 To 8
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
For i = 0 To 8
    If possible_subset(i) = True Then
        statement = statement + Str(i) + " "
    End If
Next

List1.AddItem ""
List1.AddItem "Backward Selection Finish! Best Subset: [" & "" & statement & " " & "], Goodness:" & " " & best_goodness


End Sub

Private Sub Partition_click()

' 清空 List 1 的內容
List1.Clear

' Read Txt File
List1.AddItem "Read Txt File"

in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat_x(0), feat_x(1), feat_x(2), feat_x(3), feat_x(4), feat_x(5), feat_x(6), feat_x(7), feat_x(8), feat_y
    
    For j = 0 To 8
        If j = 1 Or j = 2 Then
            attributes(j, i) = feat_x(j) * 100
        Else
            attributes(j, i) = feat_x(j)
        End If
    Next j
    attributes(9, i) = feat_y
    ' 全部列印出來看看
    ' List1.AddItem Str(attributes(0, i)) & "" & Str(attributes(1, i)) & "" & Str(attributes(2, i)) & "" & Str(attributes(3, i)) & "" & Str(attributes(4, i)) & "" & Str(attributes(5, i)) & "" & Str(attributes(6, i)) & "" & Str(attributes(7, i)) & "" & Str(attributes(8, i)) & "" & attributes(9, i)
    i = i + 1
    
Loop
Close #1

' (EWD 01) Equal-Width Discretization
List1.AddItem "(EWD 01) Equal-Width Discretization"
Dim attributes_EWD(9, 105) As Variant
Dim number_of_intervals As Integer
Dim max_num_in_col As Double
Dim min_num_in_col As Double
Dim width As Double

For i = 0 To 9

    Dim specific_col(105) As Variant
    For j = 0 To 105
        specific_col(j) = attributes(i, j)
    Next
    
    If i <> 9 Then
        number_of_intervals = 10
        max_num_in_col = find_max(specific_col)
        min_num_in_col = find_min(specific_col)
        width = (max_num_in_col - min_num_in_col) / number_of_intervals
        
        ' Decide 9 split points, so that 10 intervals will form
        split_point_1 = max_num_in_col - 1 * width
        split_point_2 = max_num_in_col - 2 * width
        split_point_3 = max_num_in_col - 3 * width
        split_point_4 = max_num_in_col - 4 * width
        split_point_5 = max_num_in_col - 5 * width
        split_point_6 = max_num_in_col - 6 * width
        split_point_7 = max_num_in_col - 7 * width
        split_point_8 = max_num_in_col - 8 * width
        split_point_9 = max_num_in_col - 9 * width
        
        ' Transfer attributes into proper interval
        Dim discrete_attr(105) As Double
        For j = 0 To 105
            discrete_attr(j) = -100
        Next
        
        For j = 0 To 105
        
            num = specific_col(j)
            
            If num >= split_point_1 Then
                discrete_attr(j) = 0
            ElseIf num >= split_point_2 And num < split_point_1 Then
                discrete_attr(j) = 1
            ElseIf num >= split_point_3 And num < split_point_2 Then
                discrete_attr(j) = 2
            ElseIf num >= split_point_4 And num < split_point_3 Then
                discrete_attr(j) = 3
            ElseIf num >= split_point_5 And num < split_point_4 Then
                discrete_attr(j) = 4
            ElseIf num >= split_point_6 And num < split_point_5 Then
                discrete_attr(j) = 5
            ElseIf num >= split_point_7 And num < split_point_6 Then
                discrete_attr(j) = 6
            ElseIf num >= split_point_8 And num < split_point_7 Then
                discrete_attr(j) = 7
            ElseIf num >= split_point_9 And num < split_point_8 Then
                discrete_attr(j) = 8
            ElseIf num < split_point_9 Then
                discrete_attr(j) = 9
            End If
        Next
        
        For j = 0 To 105
            attributes_EWD(i, j) = discrete_attr(j)
        Next
    Else
        For j = 0 To 105
            attributes_EWD(i, j) = attributes(i, j)
        Next
    End If
    
Next

' (2) Derive entropy for each column (attribute)
List1.AddItem "(EWD 02) Derive entropy for each column (attribute)"

str_list() = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "car", "fad", "mas", "gla", "con", "adi")
num_list() = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)

For j = 0 To 9
    ' 取出特定一 column
    Dim col(105) As String
    For i = 0 To 105
        col(i) = attributes_EWD(j, i)
    Next
    
    ' 計數
    Dim count(15) As Integer
    For i = 0 To 15
        count(i) = 0
    Next
    
    For i = 0 To 105
        For k = 0 To 15
            If col(i) = str_list(k) Then
                count(num_list(k)) = count(num_list(k)) + 1
            End If
        Next
    Next
    
    '機率
    Dim prob_list(15) As Double
    For i = 0 To 15
        prob_list(i) = count(i) / 106
    Next
    
    'Entropy
    Dim entropy As Double
    Dim log_val As Double
    
    entropy = 0
    For i = 0 To 15
        If prob_list(i) = 0 Then
            entropy = entropy
        Else
            log_val = Log(prob_list(i)) / Log(2)
            entropy = entropy - prob_list(i) * log_val
        End If
    Next
    
    single_entropy(j) = entropy
    
Next


' (EWD 03) Derive entropy for each two different two columns (attributes)
List1.AddItem "(EWD 03) Derive joint entropy for each two different columns (attributes)"


For j = 0 To 9
For k = 0 To 9

' 取出特定兩個 column
Dim col_A(105) As String
Dim col_B(105) As String

For i = 0 To 105
    col_A(i) = attributes_EWD(j, i)
    col_B(i) = attributes_EWD(k, i)
Next

' 計數
Dim double_count(255) As Integer

For i = 0 To 255
    double_count(i) = 0
Next

elseif_num_list() = Array(10, 11, 12, 13, 14, 15)
elseif_str_list() = Array("car", "fad", "mas", "gla", "con", "adi")

For i = 0 To 105
    For ee = 0 To 9
        ' ee 表示 0 ~ 9
        elseif_num = ee
        If Val(col_A(i)) = elseif_num And col_B(i) = "0" Then
            double_count(16 * elseif_num + 0) = double_count(16 * elseif_num + 0) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "1" Then
            double_count(16 * elseif_num + 1) = double_count(16 * elseif_num + 1) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "2" Then
            double_count(16 * elseif_num + 2) = double_count(16 * elseif_num + 2) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "3" Then
            double_count(16 * elseif_num + 3) = double_count(16 * elseif_num + 3) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "4" Then
            double_count(16 * elseif_num + 4) = double_count(16 * elseif_num + 4) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "5" Then
            double_count(16 * elseif_num + 5) = double_count(16 * elseif_num + 5) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "6" Then
            double_count(16 * elseif_num + 6) = double_count(16 * elseif_num + 6) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "7" Then
            double_count(16 * elseif_num + 7) = double_count(16 * elseif_num + 7) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "8" Then
            double_count(16 * elseif_num + 8) = double_count(16 * elseif_num + 8) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "9" Then
            double_count(16 * elseif_num + 9) = double_count(16 * elseif_num + 9) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "car" Then
            double_count(16 * elseif_num + 10) = double_count(16 * elseif_num + 10) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "fad" Then
            double_count(16 * elseif_num + 11) = double_count(16 * elseif_num + 11) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "mas" Then
            double_count(16 * elseif_num + 12) = double_count(16 * elseif_num + 12) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "gla" Then
            double_count(16 * elseif_num + 13) = double_count(16 * elseif_num + 13) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "con" Then
            double_count(16 * elseif_num + 14) = double_count(16 * elseif_num + 14) + 1
        ElseIf Val(col_A(i)) = elseif_num And col_B(i) = "adi" Then
            double_count(16 * elseif_num + 15) = double_count(16 * elseif_num + 15) + 1
        End If
    
    Next
    
    
    For ee = 0 To 5
        ' 10 ~ 15
        elseif_num = elseif_num_list(ee)
        elseif_str = elseif_str_list(ee)
        If col_A(i) = elseif_str And col_B(i) = "0" Then
            double_count(16 * elseif_num + 0) = double_count(16 * elseif_num + 0) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "1" Then
            double_count(16 * elseif_num + 1) = double_count(16 * elseif_num + 1) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "2" Then
            double_count(16 * elseif_num + 2) = double_count(16 * elseif_num + 2) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "3" Then
            double_count(16 * elseif_num + 3) = double_count(16 * elseif_num + 3) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "4" Then
            double_count(16 * elseif_num + 4) = double_count(16 * elseif_num + 4) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "5" Then
            double_count(16 * elseif_num + 5) = double_count(16 * elseif_num + 5) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "6" Then
            double_count(16 * elseif_num + 6) = double_count(16 * elseif_num + 6) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "7" Then
            double_count(16 * elseif_num + 7) = double_count(16 * elseif_num + 7) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "8" Then
            double_count(16 * elseif_num + 8) = double_count(16 * elseif_num + 8) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "9" Then
            double_count(16 * elseif_num + 9) = double_count(16 * elseif_num + 9) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "car" Then
            double_count(16 * elseif_num + 10) = double_count(16 * elseif_num + 10) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "fad" Then
            double_count(16 * elseif_num + 11) = double_count(16 * elseif_num + 11) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "mas" Then
            double_count(16 * elseif_num + 12) = double_count(16 * elseif_num + 12) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "gla" Then
            double_count(16 * elseif_num + 13) = double_count(16 * elseif_num + 13) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "con" Then
            double_count(16 * elseif_num + 14) = double_count(16 * elseif_num + 14) + 1
        ElseIf col_A(i) = elseif_str And col_B(i) = "adi" Then
            double_count(16 * elseif_num + 15) = double_count(16 * elseif_num + 15) + 1
        End If
    Next
Next

'機率
Dim double_prob_list(255) As Double

For i = 0 To 255
    double_prob_list(i) = double_count(i) / 106
Next

'Joint Entropy

entropy = 0
For i = 0 To 255
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

' (EWD 04) Derive Symmetric Uncertainty for Each two different columns (Attributes)
List1.AddItem "(EWD 04) Derive Symmetric Uncertainty for Each two different columns (Attributes)"

Dim su As Double

For i = 0 To 9
For j = 0 To 9
    
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


'(EWD 05) Forward Selection
List1.AddItem "(EWD 05) Forward Selection"

Dim must_chosen_subset(8) As Boolean
Dim possible_subset(8) As Boolean
Dim best_goodness As Double
Dim goodness_list(8) As Double
Dim exp_attr_list(8) As Integer
Dim count_same_vars As Integer
Dim chosen_subset(8) As Double
Dim good As Double
Dim statement As String
Dim idx As Integer
Dim current_best_attr As Integer
Dim current_bad_attr As Integer
Dim current_best_goodness As Double

For i = 0 To 8
    possible_subset(i) = True
    must_chosen_subset(i) = False
Next
best_goodness = 0
idx = -1

Do
    
    ' 一個 goodness 會對應到 一個 experiment 的 col，在開始前先行清空
    For i = 0 To 8
        goodness_list(i) = -1
        exp_attr_list(i) = -1
    Next

    ' 進行同樣 attr 數目的搜索
    count_same_vars = 0
    
    For j = 0 To 8
        
        If possible_subset(j) = True Then
        
            ' 每次搜索過相同 level 的 subset 時，需要清空 chosen_subset
            ' 在這邊一併加入實驗的 attr num
            For i = 0 To 8
            
                If i = j Then
                    chosen_subset(i) = i
                    exp_attr_list(count_same_vars) = i
                Else
                    chosen_subset(i) = -1 ' zero subset 的概念
                End If
            Next
            
            '把那些 最棒的 attr 加進來
            For i = 0 To 8
                If must_chosen_subset(i) = True Then
                    chosen_subset(i) = i
                End If
            Next
            
            good = get_goodness(chosen_subset)
            goodness_list(count_same_vars) = good
            count_same_vars = count_same_vars + 1
            
            statement = ""
            For i = 0 To 8
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
For i = 0 To 8
    If must_chosen_subset(i) <> False Then
        statement = statement + Str(i) + " "
    End If
Next

List1.AddItem ""
List1.AddItem "Forward Selection Finish! Best Subset: [" & "" & statement & " " & "], Goodness:" & " " & best_goodness


'(EWD 06) Backward Selection
List1.AddItem "(EWD 06) Backward Selection"

Dim must_delete_subset(8) As Boolean

For i = 0 To 8
    possible_subset(i) = True
    must_delete_subset(i) = False
Next
best_goodness = 0
idx = -1

Do
    
    ' 一個 goodness 會對應到 一個 experiment 的 col，在開始前先行清空
    For i = 0 To 8
        goodness_list(i) = -1
        exp_attr_list(i) = -1
    Next

    ' 進行同樣 attr 數目的搜索
    count_same_vars = 0
    
    For j = 0 To 8
        
        If possible_subset(j) = True Then
        
            ' 每次搜索過相同 level 的 subset 時，需要填滿 chosen_subset
            ' 在這邊一併加入實驗的 attr num
            For i = 0 To 8
            
                If i <> j Then
                    chosen_subset(i) = i
                Else
                    exp_attr_list(count_same_vars) = i
                    chosen_subset(i) = -1 ' zero subset 的概念
                End If
            Next
            
            '把那些 要拿掉的 attr 拿掉
            For i = 0 To 8
                If must_delete_subset(i) = True Then
                    chosen_subset(i) = -1
                End If
            Next
            
            good = get_goodness(chosen_subset)
            goodness_list(count_same_vars) = good
            count_same_vars = count_same_vars + 1
            
            statement = ""
            For i = 0 To 8
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
For i = 0 To 8
    If possible_subset(i) = True Then
        statement = statement + Str(i) + " "
    End If
Next

List1.AddItem ""
List1.AddItem "Backward Selection Finish! Best Subset: [" & "" & statement & " " & "], Goodness:" & " " & best_goodness

End Sub
Private Function sort_col(arr)
    
    Dim arr_sort(105) As Double
    Dim jj As Integer
    
    ' 先讓 arr_sort 裡面都是-100
    For jj = 0 To 105
        arr_sort(jj) = -100
    Next
    
    Dim max_number As Double
    Dim max_number_idx As Integer
    Dim kk As Integer
    
    Dim count As Integer
    
    For count = 0 To 105
        
        ' 找到當前 arr 裡面的最大值 以及 他的 index
        max_number = -1000
        max_number_idx = -1000
        
        For kk = 0 To 105
            If arr(kk) > max_number Then
                max_number = arr(kk)
                max_number_idx = kk
            End If
        Next
        
        ' 把它變最小，這樣之後就不會挑到他了
        arr(max_number_idx) = -1000
        
        ' 把當前最大值存到 arr_sort 裡面
        arr_sort(count) = max_number
    Next
    sort_col = arr_sort
End Function

' Helper Function: find_max
Private Function find_max(col)
    Dim max_num As Long
    Dim ii As Integer
    
    max_num = -999
    For ii = 0 To 105
        If col(ii) > max_num Then
            max_num = col(ii)
        End If
    Next
    find_max = max_num
End Function
' Helper Function: find_min
Private Function find_min(col)
    Dim min_num As Long
    Dim ii As Integer
    min_num = 10000
    For ii = 0 To 105
        If col(ii) < min_num Then
            min_num = col(ii)
        End If
    Next
    find_min = min_num
End Function

' (5.5) Define argmax
Private Function argmax(arr)
    
    Dim max_num As Double
    Dim max_num_idx As Integer
    
    max_num = -1
    max_num_idx = -1
    
    For i = 0 To 8
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
        relevant_score = relevant_score + symmetric_uncertainty(attr_num, 9)
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

