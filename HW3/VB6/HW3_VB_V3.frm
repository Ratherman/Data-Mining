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
Dim k As Single


' 讀資料相關
Dim in_file As String
Dim feat_y As String
Dim feat_x(9) As Double

' 主要儲存資訊之二維矩陣
Dim Attributes(9, 213) As Variant
Dim attributes_EWD(9, 213) As Variant
Dim attributes_EFD(9, 213) As Variant
Dim attributes_sort(9, 213) As Variant

' 用在 Equal Width Discretization
Dim number_of_intervals As Integer
Dim max_num_in_col As Double
Dim min_num_in_col As Double
Dim specific_col(213) As Variant
Dim discrete_attr(213) As Double

' 用在 Equal Frequency Discretization

Dim column(213) As Double
Dim column_sort As Variant


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

        
Private Sub Command1_Click()
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
        Attributes(j, i) = feat_x(j + 1)
    Next j
    
    Attributes(9, i) = feat_y
    i = i + 1
Loop
Close #1

For i = 0 To 9
    If i <> 9 Then
        
        For j = 0 To 213
            column(j) = Attributes(i, j)
        Next
        
        column_sort = sort_col(column)
        
        For j = 0 To 213
            attributes_sort(i, j) = column_sort(j)
        Next
        
    Else
        
        For j = 0 To 213
            attributes_sort(9, j) = Attributes(9, j)
        Next
        
    End If
Next

Dim specific_col_sort(213) As Variant
Dim specific_col(213) As Variant

For j = 0 To 9

    For i = 0 To 213
        specific_col_sort(i) = attributes_sort(j, i)
        specific_col(i) = Attributes(j, i)
    Next
    
    If j <> 9 Then
        split_point_1 = specific_col_sort(1 * 21)
        split_point_2 = specific_col_sort(2 * 21)
        split_point_3 = specific_col_sort(3 * 21)
        split_point_4 = specific_col_sort(4 * 21)
        split_point_5 = specific_col_sort(5 * 21)
        split_point_6 = specific_col_sort(6 * 21)
        split_point_7 = specific_col_sort(7 * 21)
        split_point_8 = specific_col_sort(8 * 21)
        split_point_9 = specific_col_sort(9 * 21)
    
        ' transfer attributes into proper interval
        For i = 0 To 213
            discrete_attr(i) = -100
        Next
        
        For i = 0 To 213
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
        
        Dim sp_str_1 As Double
        Dim sp_str_2 As Double
        Dim sp_str_3 As Double
        Dim sp_str_4 As Double
        Dim sp_str_5 As Double
        Dim sp_str_6 As Double
        Dim sp_str_7 As Double
        Dim sp_str_8 As Double
        Dim sp_str_9 As Double
        
        
        sp_str_1 = split_point_1
        sp_str_2 = split_point_2
        sp_str_3 = split_point_3
        sp_str_4 = split_point_4
        sp_str_5 = split_point_5
        sp_str_6 = split_point_6
        sp_str_7 = split_point_7
        sp_str_8 = split_point_8
        sp_str_9 = split_point_9
        
        List1.AddItem "第" & "" & Str(j) & "" & "個 attribute 的十個區間由大至小排序"
        List1.AddItem "  第一區間: [" & "" & Str(sp_str_1) & "" & ", " & "" & "Max num]"
        List1.AddItem "  第二區間: [" & "" & Str(sp_str_2) & "" & ", " & "" & Str(sp_str_1) & "" & ")"
        List1.AddItem "  第三區間: [" & "" & Str(sp_str_3) & "" & ", " & "" & Str(sp_str_2) & "" & ")"
        List1.AddItem "  第四區間: [" & "" & Str(sp_str_4) & "" & ", " & "" & Str(sp_str_3) & "" & ")"
        List1.AddItem "  第五區間: [" & "" & Str(sp_str_5) & "" & ", " & "" & Str(sp_str_4) & "" & ")"
        List1.AddItem "  第六區間: [" & "" & Str(sp_str_6) & "" & ", " & "" & Str(sp_str_5) & "" & ")"
        List1.AddItem "  第七區間: [" & "" & Str(sp_str_7) & "" & ", " & "" & Str(sp_str_6) & "" & ")"
        List1.AddItem "  第八區間: [" & "" & Str(sp_str_8) & "" & ", " & "" & Str(sp_str_7) & "" & ")"
        List1.AddItem "  第九區間: [" & "" & Str(sp_str_9) & "" & ", " & "" & Str(sp_str_8) & "" & ")"
        List1.AddItem "  第十區間: [" & "" & "Min num" & "" & ", " & "" & Str(sp_str_9) & "" & ")"
        List1.AddItem ""
        
    
        For i = 0 To 213
            attributes_EFD(j, i) = discrete_attr(i)
        Next
    Else
        For i = 0 To 213
            attributes_EFD(j, i) = Attributes(j, i)
        Next
    End If
Next

List1.AddItem ""
List1.AddItem "=================================="
List1.AddItem "(3) Selective Naive Bayes Forward Selection"
List1.AddItem "=================================="
List1.AddItem ""


Dim possible_subset(8) As Boolean
Dim must_chosen_subset(8) As Boolean
Dim chosen_subset(8) As Boolean

' 初始化 相關 subset
For i = 0 To 8
    possible_subset(i) = True ' 這表示所有可能的 attribute
    must_chosen_subset(i) = False ' 這用來儲存未來表現很好的 attribute
Next


' 進入迴圈
Dim current_best_accuracy As Double
Dim best_accuracy As Double
Dim acc_list(8) As Double
Dim acc As Double
Dim statement As String
Dim jj As Integer


best_accuracy = 0
Do
    For jj = 0 To 8
        acc_list(jj) = -999
    Next
    ' ========================
    ' 針對相同數目的 attr 實驗
    ' ========================
    For jj = 0 To 8
        ' ==================
        ' 清空 chosen_subset
        ' ==================
        For i = 0 To 8
            chosen_subset(i) = False
        Next
        ' =======================================================
        ' 先把 must_chosen 的 attribute 加進去 chosen_subset 裡面
        ' =======================================================
        For i = 0 To 8
            If must_chosen_subset(i) = True Then
                chosen_subset(i) = True
            End If
        Next
        
        ' =========================
        ' 現在要實驗 jj-th attribute
        ' =========================
        Dim num_of_run As Integer
        
        If possible_subset(jj) = True Then
            ' ===========================================
            ' 把 j-th attribute 加進去 chosen_subset 裡面
            ' ===========================================
            chosen_subset(jj) = True
            
            ' ==============================
            ' 使用 SNB with Laplace Estimate
            ' ==============================
            acc = selective_naive_bayes(attributes_EFD, chosen_subset)
            
            ' ==============
            ' 暫存目前的 Acc
            ' ==============
            
            acc_list(jj) = acc
        
            ' =====================
            ' 準備 print 出相關資訊
            ' =====================
            statement = ""
            For i = 0 To 8
                If chosen_subset(i) = True Then
                    statement = statement + Str(i) + " "
                End If
            Next
            
            ' ============
            ' Print 出資訊
            ' ============
            List1.AddItem "Chosen Subset: [" & "" & statement & " " & "], Accuracy:" & " " & Str(acc) & "" & "%"
            
        End If
    Next
    
    Dim max_acc_pos As Integer
    max_acc_pos = argmax(acc_list)
    current_best_accuracy = acc_list(max_acc_pos)
    
    If current_best_accuracy < best_accuracy Then
        Exit Do
    Else
        best_accuracy = current_best_accuracy
        possible_subset(max_acc_pos) = False
        must_chosen_subset(max_acc_pos) = True
    End If
    
Loop

' =====================
' 準備 print 出結果資訊
' =====================
statement = ""
For i = 0 To 8
    If must_chosen_subset(i) = True Then
        statement = statement + Str(i) + " "
    End If
Next

' ============
' Print 出資訊
' ============
List1.AddItem ""
List1.AddItem "Finished! Best Subset: [" & "" & statement & " " & "], Accuracy:" & " " & Str(best_accuracy) & "" & "%"

End Sub

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
    
        Attributes(j, i) = feat_x(j + 1)
        
    Next j
    
    Attributes(9, i) = feat_y
    i = i + 1
Loop
Close #1

List1.AddItem ""
List1.AddItem "=================================="
List1.AddItem "(2) Perform Equal Width Discretization"
List1.AddItem "=================================="
List1.AddItem ""

For i = 0 To 9
    
    For k = 0 To 213
        specific_col(k) = Attributes(i, k)
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
        
        Dim sp_str_1 As Double
        Dim sp_str_2 As Double
        Dim sp_str_3 As Double
        Dim sp_str_4 As Double
        Dim sp_str_5 As Double
        Dim sp_str_6 As Double
        Dim sp_str_7 As Double
        Dim sp_str_8 As Double
        Dim sp_str_9 As Double
        
        sp_str_1 = split_point_1
        sp_str_2 = split_point_2
        sp_str_3 = split_point_3
        sp_str_4 = split_point_4
        sp_str_5 = split_point_5
        sp_str_6 = split_point_6
        sp_str_7 = split_point_7
        sp_str_8 = split_point_8
        sp_str_9 = split_point_9
        
        List1.AddItem "第" & "" & Str(i) & "" & "個 attribute 的十個區間由大至小排序"
        List1.AddItem "  第一區間: [" & "" & Str(sp_str_1) & "" & ", " & "" & "Max num]"
        List1.AddItem "  第二區間: [" & "" & Str(sp_str_2) & "" & ", " & "" & Str(sp_str_1) & "" & ")"
        List1.AddItem "  第三區間: [" & "" & Str(sp_str_3) & "" & ", " & "" & Str(sp_str_2) & "" & ")"
        List1.AddItem "  第四區間: [" & "" & Str(sp_str_4) & "" & ", " & "" & Str(sp_str_3) & "" & ")"
        List1.AddItem "  第五區間: [" & "" & Str(sp_str_5) & "" & ", " & "" & Str(sp_str_4) & "" & ")"
        List1.AddItem "  第六區間: [" & "" & Str(sp_str_6) & "" & ", " & "" & Str(sp_str_5) & "" & ")"
        List1.AddItem "  第七區間: [" & "" & Str(sp_str_7) & "" & ", " & "" & Str(sp_str_6) & "" & ")"
        List1.AddItem "  第八區間: [" & "" & Str(sp_str_8) & "" & ", " & "" & Str(sp_str_7) & "" & ")"
        List1.AddItem "  第九區間: [" & "" & Str(sp_str_9) & "" & ", " & "" & Str(sp_str_8) & "" & ")"
        List1.AddItem "  第十區間: [" & "" & "Min num" & "" & ", " & "" & Str(sp_str_9) & "" & ")"
        List1.AddItem ""
        
        
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
            attributes_EWD(i, j) = Attributes(i, j)
        Next
    End If
    
Next
List1.AddItem ""
List1.AddItem "=================================="
List1.AddItem "(3) Selective Naive Bayes Forward Selection"
List1.AddItem "=================================="
List1.AddItem ""


Dim possible_subset(8) As Boolean
Dim must_chosen_subset(8) As Boolean
Dim chosen_subset(8) As Boolean

' 初始化 相關 subset
For i = 0 To 8
    possible_subset(i) = True ' 這表示所有可能的 attribute
    must_chosen_subset(i) = False ' 這用來儲存未來表現很好的 attribute
Next


' 進入迴圈
Dim current_best_accuracy As Double
Dim best_accuracy As Double
Dim acc_list(8) As Double
Dim acc As Double
Dim statement As String
Dim jj As Integer


best_accuracy = 0
Do
    For jj = 0 To 8
        acc_list(jj) = -999
    Next
    ' ========================
    ' 針對相同數目的 attr 實驗
    ' ========================
    For jj = 0 To 8
        ' ==================
        ' 清空 chosen_subset
        ' ==================
        For i = 0 To 8
            chosen_subset(i) = False
        Next
        ' =======================================================
        ' 先把 must_chosen 的 attribute 加進去 chosen_subset 裡面
        ' =======================================================
        For i = 0 To 8
            If must_chosen_subset(i) = True Then
                chosen_subset(i) = True
            End If
        Next
        
        ' =========================
        ' 現在要實驗 jj-th attribute
        ' =========================
        Dim num_of_run As Integer
        
        If possible_subset(jj) = True Then
            ' ===========================================
            ' 把 j-th attribute 加進去 chosen_subset 裡面
            ' ===========================================
            chosen_subset(jj) = True
            
            ' ==============================
            ' 使用 SNB with Laplace Estimate
            ' ==============================
            acc = selective_naive_bayes(attributes_EWD, chosen_subset)
            
            ' ==============
            ' 暫存目前的 Acc
            ' ==============
            
            acc_list(jj) = acc
        
            ' =====================
            ' 準備 print 出相關資訊
            ' =====================
            statement = ""
            For i = 0 To 8
                If chosen_subset(i) = True Then
                    statement = statement + Str(i) + " "
                End If
            Next
            
            ' ============
            ' Print 出資訊
            ' ============
            List1.AddItem "Chosen Subset: [" & "" & statement & " " & "], Accuracy:" & " " & Str(acc) & "" & "%"
            
        End If
    Next
    
    Dim max_acc_pos As Integer
    max_acc_pos = argmax(acc_list)
    current_best_accuracy = acc_list(max_acc_pos)
    
    If current_best_accuracy < best_accuracy Then
        Exit Do
    Else
        best_accuracy = current_best_accuracy
        possible_subset(max_acc_pos) = False
        must_chosen_subset(max_acc_pos) = True
    End If
    
Loop

' =====================
' 準備 print 出結果資訊
' =====================
statement = ""
For i = 0 To 8
    If must_chosen_subset(i) = True Then
        statement = statement + Str(i) + " "
    End If
Next

' ============
' Print 出資訊
' ============
List1.AddItem ""
List1.AddItem "Finished! Best Subset: [" & "" & statement & " " & "], Accuracy:" & " " & Str(best_accuracy) & "" & "%"



End Sub
' Core Function: Perform SNB Algorithm with Laplace Estimate
Private Function selective_naive_bayes(Attributes, col_boolean):
    
    ' 準備好 Y_col
    Dim Y_col(213) As Double
    For i = 0 To 213
        Y_col(i) = Attributes(9, i)
    Next
    
    
    ' 準備好 Datum，先把他初始化成 -999
    Dim Datum(8, 213) As Double
    For i = 0 To 8
        For j = 0 To 213
            Datum(i, j) = -999
        Next
    Next
    
    ' 然後把需要使用的資料(每一次都是不一樣的column)倒進 Datum 裡面
    For i = 0 To 8
        If col_boolean(i) = True Then
            For j = 0 To 213
                Datum(i, j) = Attributes(i, j)
            Next
        End If
    Next
    
    ' 等會兒要記數
    Dim ACC_Count As Double
    ACC_Count = 0
    
    Dim ii As Integer
    
    
    For ii = 0 To 213
    
        ' ===============================================
        ' Step 1: 取出 測試資料，包含 X_test(8) 和 Y_test
        ' ===============================================
        Dim X_test(8) As Double
        Dim Y_test As Double
        
        For j = 0 To 8
            X_test(j) = Datum(j, ii) '這214 筆資料，每一筆都會當一次測試資料
        Next
        Y_test = Y_col(ii)
        
        ' ===========================================================
        ' Step 2: 取出 訓練資料，包含 X_train(212, 9) 和 Y_train(212)
        ' ===========================================================
        Dim X_train(8, 212) As Double
        Dim Y_train(212) As Double
        Dim train_counter As Integer
        
        For j = 0 To 8
            train_counter = 0 ' 每個 column 都需要歸零一次
            If col_boolean(j) = True Then
                
                For k = 0 To 213
                    If k = ii Then ' 表示這筆資料是 testing 的
                        train_counter = train_counter
                    Else
                        X_train(j, train_counter) = Datum(j, k)
                        Y_train(train_counter) = Y_col(k)
                        train_counter = train_counter + 1
                    End If
                Next
            End If
        Next
        
        ' ======================================
        ' Step 3: 用 test 和 train 衡量 Accuracy
        ' ======================================
        
        ' ==================
        ' Step 3.1 計算 p(c)
        ' ==================
        Dim count_c1, count_c2, count_c3, count_c5, count_c6, count_c7 As Double
        Dim prob_c1, prob_c2, prob_c3, prob_c5, prob_c6, prob_c7 As Double
        
        count_c1 = 0
        count_c2 = 0
        count_c3 = 0
        count_c5 = 0
        count_c6 = 0
        count_c7 = 0
        
        For j = 0 To 212
            If (Y_train(j) = 1) Then
                count_c1 = count_c1 + 1
            ElseIf (Y_train(j) = 2) Then
                count_c2 = count_c2 + 1
            ElseIf (Y_train(j) = 3) Then
                count_c3 = count_c3 + 1
            ElseIf (Y_train(j) = 5) Then
                count_c5 = count_c5 + 1
            ElseIf (Y_train(j) = 6) Then
                count_c6 = count_c6 + 1
            ElseIf (Y_train(j) = 7) Then
                count_c7 = count_c7 + 1
            End If
        Next
        
        prob_c1 = count_c1 / 213
        prob_c2 = count_c2 / 213
        prob_c3 = count_c3 / 213
        prob_c5 = count_c5 / 213
        prob_c6 = count_c6 / 213
        prob_c7 = count_c7 / 213
                
        ' ====================
        ' Step 3.2 計算 p(x,c)
        ' ====================
        Dim cond_prob_c1, cond_prob_c2, cond_prob_c3, cond_prob_c5, cond_prob_c6, cond_prob_c7 As Double
        Dim target_count_c1, target_count_c2, target_count_c3, target_count_c5, target_count_c6, target_count_c7 As Double
        Dim target_val As Double
        Dim target_column(212) As Double ' 這個會在 laplace estimate 用到
        
        cond_prob_c1 = 1
        cond_prob_c2 = 1
        cond_prob_c3 = 1
        cond_prob_c5 = 1
        cond_prob_c6 = 1
        cond_prob_c7 = 1
        
        ' loop 過每個 column
        For k = 0 To 8
        
            ' 只有被選中的 column 能夠進來
            If col_boolean(k) = True Then
                
                For j = 0 To 212
                    target_column(j) = X_train(k, j)
                Next
                
                target_val = X_test(k)
                
                target_count_c1 = 0
                target_count_c2 = 0
                target_count_c3 = 0
                target_count_c5 = 0
                target_count_c6 = 0
                target_count_c7 = 0
            
                For j = 0 To 212
                    If Y_train(j) = 1 And target_val = target_column(j) Then
                        target_count_c1 = target_count_c1 + 1
                    ElseIf Y_train(j) = 2 And target_val = target_column(j) Then
                        target_count_c2 = target_count_c2 + 1
                    ElseIf Y_train(j) = 3 And target_val = target_column(j) Then
                        target_count_c3 = target_count_c3 + 1
                    ElseIf Y_train(j) = 5 And target_val = target_column(j) Then
                        target_count_c5 = target_count_c5 + 1
                    ElseIf Y_train(j) = 6 And target_val = target_column(j) Then
                        target_count_c6 = target_count_c6 + 1
                    ElseIf Y_train(j) = 7 And target_val = target_column(j) Then
                        target_count_c7 = target_count_c7 + 1
                    End If
                Next

                cond_prob_c1 = cond_prob_c1 * (target_count_c1 + 1) / (count_c1 + unique_type_count(target_column))
                cond_prob_c2 = cond_prob_c2 * (target_count_c2 + 1) / (count_c2 + unique_type_count(target_column))
                cond_prob_c3 = cond_prob_c3 * (target_count_c3 + 1) / (count_c3 + unique_type_count(target_column))
                cond_prob_c5 = cond_prob_c5 * (target_count_c5 + 1) / (count_c5 + unique_type_count(target_column))
                cond_prob_c6 = cond_prob_c6 * (target_count_c6 + 1) / (count_c6 + unique_type_count(target_column))
                cond_prob_c7 = cond_prob_c7 * (target_count_c7 + 1) / (count_c7 + unique_type_count(target_column))
            End If
            
        Next
        ' ===================
        ' Step 3.3 計算 score
        ' ===================
        Dim score_list(5) As Double
        score_list(0) = prob_c1 * cond_prob_c1
        score_list(1) = prob_c2 * cond_prob_c2
        score_list(2) = prob_c3 * cond_prob_c3
        score_list(3) = prob_c5 * cond_prob_c5
        score_list(4) = prob_c6 * cond_prob_c6
        score_list(5) = prob_c7 * cond_prob_c7
        
        ' =====================================================
        ' Step 3.4 比較結果是否與 Y_test 一致，有的話 count + 1
        ' =====================================================
        Dim max_position As Integer
        max_position = find_max_pos(score_list)
        
        If (Y_test = 1) And (max_position = 0) Then
            ACC_Count = ACC_Count + 1
        ElseIf (Y_test = 2) And (max_position = 1) Then
            ACC_Count = ACC_Count + 1
        ElseIf (Y_test = 3) And (max_position = 2) Then
            ACC_Count = ACC_Count + 1
        ElseIf (Y_test = 5) And (max_position = 3) Then
            ACC_Count = ACC_Count + 1
        ElseIf (Y_test = 6) And (max_position = 4) Then
            ACC_Count = ACC_Count + 1
        ElseIf (Y_test = 7) And (max_position = 5) Then
            ACC_Count = ACC_Count + 1
        Else
            ACC_Count = ACC_Count
        End If
        
    Next

    selective_naive_bayes = ACC_Count / 214 * 100
End Function

' Helper Function: count number of unique types in a column
' Laplace Estimate 需要這個數值(用在分母)
' 我可以用這種方式去找出 unique 的數量是因為我已經有把所有的可能值換成 0 ~ 9 其中一個數字了
Private Function unique_type_count(col_array):
    Dim toggle_0 As Integer
    Dim toggle_1 As Integer
    Dim toggle_2 As Integer
    Dim toggle_3 As Integer
    Dim toggle_4 As Integer
    Dim toggle_5 As Integer
    Dim toggle_6 As Integer
    Dim toggle_7 As Integer
    Dim toggle_8 As Integer
    Dim toggle_9 As Integer
    Dim num_unique As Integer
    
    toggle_0 = 0
    toggle_1 = 0
    toggle_2 = 0
    toggle_3 = 0
    toggle_4 = 0
    toggle_5 = 0
    toggle_6 = 0
    toggle_7 = 0
    toggle_8 = 0
    toggle_9 = 0
    
    ' 這裡只有用 213 筆資料是因為我有做 dataset 的切割
    For i = 0 To 212
        If col_array(i) = 0 Then
            toggle_0 = 1
        End If
        
        If col_array(i) = 1 Then
            toggle_1 = 1
        End If
        
        If col_array(i) = 2 Then
            toggle_2 = 1
        End If
        
        If col_array(i) = 3 Then
            toggle_3 = 1
        End If
        
        If col_array(i) = 4 Then
            toggle_4 = 1
        End If
        
        If col_array(i) = 5 Then
            toggle_5 = 1
        End If
        
        If col_array(i) = 6 Then
            toggle_6 = 1
        End If
        
        If col_array(i) = 7 Then
            toggle_7 = 1
        End If
        
        If col_array(i) = 8 Then
            toggle_8 = 1
        End If
        
        If col_array(i) = 9 Then
            toggle_9 = 1
        End If
    Next
    
    num_unique = toggle_0 + toggle_1 + toggle_2 + toggle_3 + toggle_4 + toggle_5 + toggle_6 + toggle_7 + toggle_8 + toggle_9
    unique_type_count = num_unique
End Function

' Helper Function: sort col
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
    Dim max_num As Double
    Dim ii As Integer
    
    max_num = -999
    For ii = 0 To 213
        If col(ii) > max_num Then
            max_num = col(ii)
        End If
    Next
    find_max = max_num
End Function

' Helper Function: find_max_pos
Private Function find_max_pos(col)
    Dim max_num As Double
    Dim max_idx As Integer
    Dim ii As Integer
    
    max_num = -999
    
    For ii = 0 To 5
        If col(ii) > max_num Then
            max_num = col(ii)
            max_idx = ii
        End If
    Next
    find_max_pos = max_idx
End Function



' Helper Function: find_min
Private Function find_min(col)
    Dim min_num As Double
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
    
    max_num = -999
    max_num_idx = -999
    
    For i = 0 To 8
        If arr(i) > max_num Then
            max_num = arr(i)
            max_num_idx = i
        End If
    Next
    
    argmax = max_num_idx
    
End Function

