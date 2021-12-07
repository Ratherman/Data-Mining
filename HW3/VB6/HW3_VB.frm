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
      Width           =   16095
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
      Text            =   "glass.txt"
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

' 讀資料相關
Dim in_file As String
Dim feat_y As String
Dim feat_x(9) As Double

' 主要儲存資訊之二維矩陣
Dim attributes(9, 214) As Variant
Dim attributes_EWD(9, 214) As Variant

' 用在 Equal Width Discretization
Dim number_of_intervals As Integer
Dim max_num_in_col As Double
Dim min_num_in_col As Double
Dim specific_col(214) As Variant
Dim discrete_attr(214) As Double

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

        
Private Sub Partition_click()

' 清空 List 1 的內容
List1.Clear

List1.AddItem "=================================="
List1.AddItem "(1) Read Data"
List1.AddItem "=================================="

in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat_x(0), feat_x(1), feat_x(2), feat_x(3), feat_x(4), feat_x(5), feat_x(6), feat_x(7), feat_x(8), feat_x(9), feat_y
    
    For j = 0 To 8
    
        If j = 1 Then
            attributes(j, i) = feat_x(j) * 1000 ' scale 太小，不好 discretization
        Else
            attributes(j, i) = feat_x(j)
        End If
        
'        attributes(j, i) = feat_x(j + 1)
    Next j
    
    attributes(9, i) = feat_y
    ' 全部列印出來看看
    ' List1.AddItem Str(attributes(0, i)) & " " & Str(attributes(1, i)) & " " & Str(attributes(2, i)) & " " & Str(attributes(3, i)) & " " & Str(attributes(4, i)) & " " & Str(attributes(5, i)) & " " & Str(attributes(6, i)) & " " & Str(attributes(7, i)) & " " & Str(attributes(8, i)) & " " & attributes(9, i)
    i = i + 1
Loop
Close #1

List1.AddItem ""
List1.AddItem "=================================="
List1.AddItem "(2) Perform Equal Width Discretization"
List1.AddItem "=================================="
List1.AddItem ""

For i = 0 To 9
    
    For j = 0 To 213
        specific_col(j) = attributes(i, j)
    Next
    
    If i <> 9 Then
        number_of_intervals = 10
        max_num_in_col = find_max(specific_col)
        min_num_in_col = find_min(specific_col)
        
        Dim width As Double
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
        
        
        If i = 1 Then
            split_point_1 = split_point_1 / 1000
            split_point_2 = split_point_2 / 1000
            split_point_3 = split_point_3 / 1000
            split_point_4 = split_point_4 / 1000
            split_point_5 = split_point_5 / 1000
            split_point_6 = split_point_6 / 1000
            split_point_7 = split_point_7 / 1000
            split_point_8 = split_point_8 / 1000
            split_point_9 = split_point_9 / 1000
        
            List1.AddItem "第" & "" & Str(i) & "" & "個 attribute 的十個區間由大至小排序"
            List1.AddItem "  第一區間: [" & "" & Str(split_point_1) & "" & ", " & "" & "Max num]"
            List1.AddItem "  第二區間: [" & "" & Str(split_point_2) & "" & ", " & "" & Str(split_point_1) & "" & ")"
            List1.AddItem "  第三區間: [" & "" & Str(split_point_3) & "" & ", " & "" & Str(split_point_2) & "" & ")"
            List1.AddItem "  第四區間: [" & "" & Str(split_point_4) & "" & ", " & "" & Str(split_point_3) & "" & ")"
            List1.AddItem "  第五區間: [" & "" & Str(split_point_5) & "" & ", " & "" & Str(split_point_4) & "" & ")"
            List1.AddItem "  第六區間: [" & "" & Str(split_point_6) & "" & ", " & "" & Str(split_point_5) & "" & ")"
            List1.AddItem "  第七區間: [" & "" & Str(split_point_7) & "" & ", " & "" & Str(split_point_6) & "" & ")"
            List1.AddItem "  第八區間: [" & "" & Str(split_point_8) & "" & ", " & "" & Str(split_point_7) & "" & ")"
            List1.AddItem "  第九區間: [" & "" & Str(split_point_9) & "" & ", " & "" & Str(split_point_8) & "" & ")"
            List1.AddItem "  第十區間: [" & "" & "Min num" & "" & ", " & "" & Str(split_point_9) & "" & ")"
            List1.AddItem ""

        
        Else
        
            'List1.AddItem "第" & "" & Str(i) & "" & "個 attribute 的十個區間由大至小排序"
            'List1.AddItem "  第一區間: [" & "" & Str(split_point_1) & "" & ", " & "" & "Max num]"
            'List1.AddItem "  第二區間: [" & "" & Str(split_point_2) & "" & ", " & "" & Str(split_point_1) & "" & ")"
            'List1.AddItem "  第三區間: [" & "" & Str(split_point_3) & "" & ", " & "" & Str(split_point_2) & "" & ")"
            'List1.AddItem "  第四區間: [" & "" & Str(split_point_4) & "" & ", " & "" & Str(split_point_3) & "" & ")"
            'List1.AddItem "  第五區間: [" & "" & Str(split_point_5) & "" & ", " & "" & Str(split_point_4) & "" & ")"
            'List1.AddItem "  第六區間: [" & "" & Str(split_point_6) & "" & ", " & "" & Str(split_point_5) & "" & ")"
            'List1.AddItem "  第七區間: [" & "" & Str(split_point_7) & "" & ", " & "" & Str(split_point_6) & "" & ")"
            'List1.AddItem "  第八區間: [" & "" & Str(split_point_8) & "" & ", " & "" & Str(split_point_7) & "" & ")"
            'List1.AddItem "  第九區間: [" & "" & Str(split_point_9) & "" & ", " & "" & Str(split_point_8) & "" & ")"
            'List1.AddItem "  第十區間: [" & "" & "Min num" & "" & ", " & "" & Str(split_point_9) & "" & ")"
            'List1.AddItem ""
        
        End If
        
        ' Transfer attributes into proper interval
        For j = 0 To 213
            discrete_attr(j) = -100
        Next
        
        For j = 0 To 213
        
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
        
        For j = 0 To 213
            attributes_EWD(i, j) = discrete_attr(j)
        Next
    Else
        For j = 0 To 213
            attributes_EWD(i, j) = attributes(i, j)
        Next
    End If
    
Next
List1.AddItem "等一下記得要打開"

List1.AddItem ""
List1.AddItem "=================================="
List1.AddItem "(3) Use Selective Naive Bayes (with Laplace Estimate)"
List1.AddItem "=================================="
List1.AddItem ""

Dim must_chosen_subset(8) As Boolean
Dim possible_subset(8) As Boolean
Dim best_accuracy As Double
Dim accuracy_list(8) As Double
Dim exp_attr_list(8) As Integer
Dim count_same_vars As Integer
Dim chosen_subset(8) As Double
Dim acc As Double
Dim statement As String
Dim idx As Integer
Dim current_best_attr As Integer
Dim current_bad_attr As Integer
Dim current_best_accuracy As Double

'For i = 0 To 8
'    possible_subset(i) = True
'    must_chosen_subset(i) = False
'Next
'best_accuracy = 0
'idx = -1
'
'Do
'
'    ' 一個 accuracy 會對應到 一個 experiment 的 col，在開始前先行清空
'    For i = 0 To 8
'        accuracy_list(i) = -1
'        exp_attr_list(i) = -1
'    Next
'
'    ' 進行同樣 attr 數目的搜索
'    count_same_vars = 0
'
'    For j = 0 To 8
'
'        If possible_subset(j) = True Then
'
'            ' 每次搜索過相同 level 的 subset 時，需要清空 chosen_subset
'            ' 在這邊一併加入實驗的 attr num
'            For i = 0 To 8
'
'                If i = j Then
'                    chosen_subset(i) = i
'                    exp_attr_list(count_same_vars) = i
'                Else
'                    chosen_subset(i) = -1 ' zero subset 的概念
'                End If
'            Next
'
'            '把那些 最棒的 attr 加進來
'            For i = 0 To 8
'                If must_chosen_subset(i) = True Then
'                    chosen_subset(i) = i
'                End If
'            Next
'
'            acc = get_accuracy(chosen_subset)
'            accuracy_list(count_same_vars) = acc
'            count_same_vars = count_same_vars + 1
'
'            statement = ""
'            For i = 0 To 8
'                If chosen_subset(i) <> -1 Then
'                    statement = statement + Str(i) + " "
'                End If
'            Next
'            List1.AddItem "Chosen Subset: [" & "" & statement & " " & "], Accuracy:" & " " & acc
'        End If
'    Next
'
'    ' 找到 accuracy　最大的 idx，並透過 exp_attr_list 找到 attr_num 和 goodness
'    idx = argmax(accuracy_list)
'    current_best_attr = exp_attr_list(idx)
'    current_best_accuracy = accuracy_list(idx)
'
'    '設定停止條件
'    ' (1) current_best_accuracy 比 best_accuracy 低
'    If current_best_accuracy < best_accuracy Then
'        Exit Do
'    Else
'        must_chosen_subset(current_best_attr) = True
'        possible_subset(current_best_attr) = False
'        best_accuracy = current_best_accuracy
'    End If
'Loop

'statement = ""
'For i = 0 To 8
'    If must_chosen_subset(i) <> False Then
'        statement = statement + Str(i) + " "
'    End If
'Next

'List1.AddItem ""
'List1.AddItem "Forward Selection Finish! Best Subset: [" & "" & statement & " " & "], Accuracy:" & " " & best_accuracy

End Sub

Private Function sort_col(arr)
    
    Dim arr_sort(213) As Double
    Dim jj As Integer
    
    ' 先讓 arr_sort 裡面都是-100
    For jj = 0 To 213
        arr_sort(jj) = -100
    Next
    
    Dim max_number As Double
    Dim max_number_idx As Integer
    Dim kk As Integer
    
    Dim count As Integer
    
    For count = 0 To 213
        
        ' 找到當前 arr 裡面的最大值 以及 他的 index
        max_number = -1000
        max_number_idx = -1000
        
        For kk = 0 To 213
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
    For ii = 0 To 213
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
    For ii = 0 To 213
        If col(ii) < min_num Then
            min_num = col(ii)
        End If
    Next
    find_min = min_num
End Function
' Helper Function: argmax
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

