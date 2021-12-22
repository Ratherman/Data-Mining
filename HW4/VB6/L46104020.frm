VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   12270
   Begin VB.CommandButton Command3 
      Caption         =   "Model 4: KNN + Weighted Vote"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Model 3: KNN + Majority Vote"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Model 2: Equal-Frequency + NBC"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   4335
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
      Width           =   12015
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
      Text            =   "pima.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Model 1: Equal-Width + NBC"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   4335
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

' iterator
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim col_num As Integer

' 讀資料相關
Dim in_file As String
Dim feat_y As String
Dim feat_x(7) As Double

' 主要儲存資訊之二維矩陣
Dim Attributes(8, 767) As Variant
Dim Attributes_EWD(8, 767) As Variant
Dim Attributes_EFD(8, 767) As Variant
Dim Attributes_Sort(8, 767) As Variant
Dim Attributes_Norm(8, 767) As Variant
Dim Shuffle_Attributes(8, 767) As Variant
Dim AF_1(8, 152) As Variant
Dim AF_2(8, 152) As Variant
Dim AF_3(8, 152) As Variant
Dim AF_4(8, 152) As Variant
Dim AF_5(8, 155) As Variant
Dim train_dataset_type_1 As Variant ' (8, 614)
Dim train_dataset_type_2 As Variant ' (8, 614)
Dim train_dataset_type_3 As Variant ' (8, 614)
Dim train_dataset_type_4 As Variant ' (8, 614)
Dim train_dataset_type_5 As Variant ' (8, 612)

' 主要儲存資訊的一維陣列
Dim specific_col(767) As Variant
Dim specific_col_sort(767) As Variant
Dim discrete_attr(767) As Variant
Dim random_list(767) As Variant
Dim column_sort As Variant ' 心裡知道他就是 767
Dim column(767) As Variant
Dim arr_ori_idx As Variant ' 心裡知道他就是 767
Dim specific_col_0(767) As Variant
Dim specific_col_1(767) As Variant
Dim specific_col_2(767) As Variant
Dim specific_col_3(767) As Variant
Dim specific_col_4(767) As Variant
Dim specific_col_5(767) As Variant
Dim specific_col_6(767) As Variant
Dim specific_col_7(767) As Variant
Dim specific_col_0_norm(767) As Variant
Dim specific_col_1_norm(767) As Variant
Dim specific_col_2_norm(767) As Variant
Dim specific_col_3_norm(767) As Variant
Dim specific_col_4_norm(767) As Variant
Dim specific_col_5_norm(767) As Variant
Dim specific_col_6_norm(767) As Variant
Dim specific_col_7_norm(767) As Variant
' fold correct number
Dim fold_1_correct As Double
Dim fold_2_correct As Double
Dim fold_3_correct As Double
Dim fold_4_correct As Double
Dim fold_5_correct As Double

Dim P1_BAR As Double
Dim P2_BAR As Double
Dim P3_BAR As Double
Dim P4_BAR As Double
Dim P_AVG As Double
Dim Z As Double
Dim max_attr As Double
Dim min_attr As Double





Private Sub Partition_click()
List1.Clear
List1.AddItem ""
List1.AddItem "==============================================================="
List1.AddItem "Model 1: Equal-Width Discretization + Naive Bayesian Classifier"
List1.AddItem "==============================================================="
List1.AddItem ""

' ====================
' Step 1: Read Data
' ====================
in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat_x(0), feat_x(1), feat_x(2), feat_x(3), feat_x(4), feat_x(5), feat_x(6), feat_x(7), feat_y
    
    For j = 0 To 7
        Attributes(j, i) = feat_x(j)
    Next j
    
    Attributes(8, i) = feat_y
    i = i + 1
Loop
Close #1

' Checker
' For i = 0 To 767
'    List1.AddItem Str(Attributes(0, i)) & " " & Str(Attributes(1, i)) & " " & Str(Attributes(2, i)) & " " & Str(Attributes(3, i)) & " " & Str(Attributes(4, i)) & " " & Str(Attributes(5, i)) & " " & Str(Attributes(6, i)) & " " & Str(Attributes(7, i)) & " " & Str(Attributes(8, i))
'Next

' ==================================
' Step 2: Equal-Width with 10 Bins
' ==================================
' 主要放置 split point 相關之參數
Dim split_point_1 As Double
Dim split_point_2 As Double
Dim split_point_3 As Double
Dim split_point_4 As Double
Dim split_point_5 As Double
Dim split_point_6 As Double
Dim split_point_7 As Double
Dim split_point_8 As Double
Dim split_point_9 As Double

Dim sp_str_1 As Double
Dim sp_str_2 As Double
Dim sp_str_3 As Double
Dim sp_str_4 As Double
Dim sp_str_5 As Double
Dim sp_str_6 As Double
Dim sp_str_7 As Double
Dim sp_str_8 As Double
Dim sp_str_9 As Double

Dim number_of_intervals As Integer
Dim width As Double
Dim num As Double
Dim max_num_in_col As Double
Dim min_num_in_col As Double

For i = 0 To 8
    
    For k = 0 To 767
        specific_col(k) = Attributes(i, k)
    Next
    
    If i <> 8 Then
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
        For j = 0 To 767
            discrete_attr(j) = -100
        Next
        
        For j = 0 To 767
        
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
        
        For j = 0 To 767
            Attributes_EWD(i, j) = discrete_attr(j)
        Next
    Else
        For j = 0 To 767
            Attributes_EWD(i, j) = Attributes(i, j)
        Next
    End If
Next

'Checker
'For i = 0 To 767
'    List1.AddItem Str(Attributes_EWD(0, i)) & " " & Str(Attributes_EWD(1, i)) & " " & Str(Attributes_EWD(2, i)) & " " & Str(Attributes_EWD(3, i)) & " " & Str(Attributes_EWD(4, i)) & " " & Str(Attributes_EWD(5, i)) & " " & Str(Attributes_EWD(6, i)) & " " & Str(Attributes_EWD(7, i)) & " " & Str(Attributes_EWD(8, i))
'Next

' ======================
' Step 3: Random Shuffle
' ======================
For i = 0 To 767
    random_list(i) = Rnd()
Next
arr_ori_idx = sort_col_get_idx(random_list)

'Checker
'For i = 0 To 767
'    List1.AddItem arr_ori_idx(i)
'Next
Dim ori_idx As Integer
Dim counter As Integer

counter = 0
For ii = 0 To 767
    ori_idx = arr_ori_idx(counter)
    For jj = 0 To 8
        Shuffle_Attributes(jj, counter) = Attributes_EWD(jj, ii)
    Next
    counter = counter + 1
Next

'Checker
'For i = 0 To 767
'    List1.AddItem Str(Shuffle_Attributes(0, i)) & " " & Str(Shuffle_Attributes(1, i)) & " " & Str(Shuffle_Attributes(2, i)) & " " & Str(Shuffle_Attributes(3, i)) & " " & Str(Shuffle_Attributes(4, i)) & " " & Str(Shuffle_Attributes(5, i)) & " " & Str(Shuffle_Attributes(6, i)) & " " & Str(Shuffle_Attributes(7, i)) & " " & Str(Shuffle_Attributes(8, i))
'Next

' ==========================
' Step 4: Split into 5 Folds
' ==========================

Dim counter_k1 As Integer
Dim counter_k2 As Integer
Dim counter_k3 As Integer
Dim counter_k4 As Integer
Dim counter_k5 As Integer

counter_k1 = 0
counter_k2 = 0
counter_k3 = 0
counter_k4 = 0
counter_k5 = 0

For i = 0 To 767
    If i < 153 * 1 Then
        For j = 0 To 8
            AF_1(j, counter_k1) = Shuffle_Attributes(j, i)
        Next
        counter_k1 = counter_k1 + 1
        
    ElseIf i < 153 * 2 And i >= 153 * 1 Then
        For j = 0 To 8
            AF_2(j, counter_k2) = Shuffle_Attributes(j, i)
        Next
        counter_k2 = counter_k2 + 1
    
    ElseIf i < 153 * 3 And i >= 153 * 2 Then
        For j = 0 To 8
            AF_3(j, counter_k3) = Shuffle_Attributes(j, i)
        Next
        counter_k3 = counter_k3 + 1
    
    ElseIf i < 153 * 4 And i >= 153 * 3 Then
        For j = 0 To 8
            AF_4(j, counter_k4) = Shuffle_Attributes(j, i)
        Next
        counter_k4 = counter_k4 + 1
    
    ElseIf i >= 153 * 4 Then
        For j = 0 To 8
            AF_5(j, counter_k5) = Shuffle_Attributes(j, i)
        Next
        counter_k5 = counter_k5 + 1
    End If
Next

'Checker
'For i = 0 To 152
'For i = 0 To 155
    'List1.AddItem Str(AF_1(0, i)) & " " & Str(AF_1(1, i)) & " " & Str(AF_1(2, i)) & " " & Str(AF_1(3, i)) & " " & Str(AF_1(4, i)) & " " & Str(AF_1(5, i)) & " " & Str(AF_1(6, i)) & " " & Str(AF_1(7, i)) & " " & Str(AF_1(8, i))
    'List1.AddItem Str(AF_2(0, i)) & " " & Str(AF_2(1, i)) & " " & Str(AF_2(2, i)) & " " & Str(AF_2(3, i)) & " " & Str(AF_2(4, i)) & " " & Str(AF_2(5, i)) & " " & Str(AF_2(6, i)) & " " & Str(AF_2(7, i)) & " " & Str(AF_2(8, i))
    'List1.AddItem Str(AF_5(0, i)) & " " & Str(AF_5(1, i)) & " " & Str(AF_5(2, i)) & " " & Str(AF_5(3, i)) & " " & Str(AF_5(5, i)) & " " & Str(AF_5(5, i)) & " " & Str(AF_5(6, i)) & " " & Str(AF_5(7, i)) & " " & Str(AF_5(8, i))
'Next

' ================================
' Step 5: Naive Baysian Classifier
' ================================
train_dataset_type_1 = map_into_train_dataset(153, 153, 153, 156, AF_2, AF_3, AF_4, AF_5)
train_dataset_type_2 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_3, AF_4, AF_5)
train_dataset_type_3 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_2, AF_4, AF_5)
train_dataset_type_4 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_2, AF_3, AF_5)
train_dataset_type_5 = map_into_train_dataset(153, 153, 153, 153, AF_1, AF_2, AF_3, AF_4)

fold_1_correct = naive_bayes_classifier(AF_1, 153, train_dataset_type_1, 615)
fold_2_correct = naive_bayes_classifier(AF_2, 153, train_dataset_type_2, 615)
fold_3_correct = naive_bayes_classifier(AF_3, 153, train_dataset_type_3, 615)
fold_4_correct = naive_bayes_classifier(AF_4, 153, train_dataset_type_4, 615)
fold_5_correct = naive_bayes_classifier(AF_5, 156, train_dataset_type_5, 612)

'Checker
'List1.AddItem Str(fold_1_correct) & " " & Str(fold_2_correct) & " " & Str(fold_3_correct) & " " & Str(fold_4_correct) & " " & Str(fold_5_correct)

' ==================================
' Step 6: Calculate Average Accuracy
' ==================================
P1_BAR = (fold_1_correct + fold_2_correct + fold_3_correct + fold_4_correct + fold_5_correct) / 768
List1.AddItem "[Model 1: Equal-width Discretization + Naive Bayes Classifier] Average is" & " " & Str(P1_BAR)

' ===========================================================================================================

List1.AddItem ""
List1.AddItem "==============================================================="
List1.AddItem "Model 2: Equal-Frequency Discretization + Naive Bayesian Classifier"
List1.AddItem "==============================================================="
List1.AddItem ""

' ====================
' Step 1: Read Data
' ====================
in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat_x(0), feat_x(1), feat_x(2), feat_x(3), feat_x(4), feat_x(5), feat_x(6), feat_x(7), feat_y
    
    For j = 0 To 7
        Attributes(j, i) = feat_x(j)
    Next j
    
    Attributes(8, i) = feat_y
    i = i + 1
Loop
Close #1

' ========================
' Step 2: Sort Attributes
' ========================
For i = 0 To 8
    If i <> 8 Then
        
        For j = 0 To 767
            column(j) = Attributes(i, j)
        Next
        
        column_sort = sort_col(column)
        
        For j = 0 To 767
            Attributes_Sort(i, j) = column_sort(j)
        Next
        
    Else
        
        For j = 0 To 767
            Attributes_Sort(8, j) = Attributes(8, j)
        Next
        
    End If
Next

' ====================================
' Step 3: Equal-Frequency with 10 Bins
' ====================================
For j = 0 To 8

    For i = 0 To 767
        specific_col_sort(i) = Attributes_Sort(j, i)
        specific_col(i) = Attributes(j, i)
    Next
    
    If j <> 8 Then
        split_point_1 = specific_col_sort(1 * 76)
        split_point_2 = specific_col_sort(2 * 76)
        split_point_3 = specific_col_sort(3 * 76)
        split_point_4 = specific_col_sort(4 * 76)
        split_point_5 = specific_col_sort(5 * 76)
        split_point_6 = specific_col_sort(6 * 76)
        split_point_7 = specific_col_sort(7 * 76)
        split_point_8 = specific_col_sort(8 * 76)
        split_point_9 = specific_col_sort(9 * 76)
    
        ' transfer attributes into proper interval
        For i = 0 To 767
            discrete_attr(i) = -100
        Next
        
        For i = 0 To 767
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
        
    
        For i = 0 To 767
            Attributes_EFD(j, i) = discrete_attr(i)
        Next
    Else
        For i = 0 To 767
            Attributes_EFD(j, i) = Attributes(j, i)
        Next
    End If
Next

' ======================
' Step 4: Random Shuffle
' ======================
'For i = 0 To 767
'    random_list(i) = Rnd()
'Next
'arr_ori_idx = sort_col_get_idx(random_list)

'Checker
'For i = 0 To 767
'    List1.AddItem arr_ori_idx(i)
'Next

counter = 0
For ii = 0 To 767
    ori_idx = arr_ori_idx(counter)
    For jj = 0 To 8
        Shuffle_Attributes(jj, counter) = Attributes_EFD(jj, ii)
    Next
    counter = counter + 1
Next

' ==========================
' Step 5: Split into 5 folds
' ==========================
counter_k1 = 0
counter_k2 = 0
counter_k3 = 0
counter_k4 = 0
counter_k5 = 0

For i = 0 To 767
    If i < 153 * 1 Then
        For j = 0 To 8
            AF_1(j, counter_k1) = Shuffle_Attributes(j, i)
        Next
        counter_k1 = counter_k1 + 1
        
    ElseIf i < 153 * 2 And i >= 153 * 1 Then
        For j = 0 To 8
            AF_2(j, counter_k2) = Shuffle_Attributes(j, i)
        Next
        counter_k2 = counter_k2 + 1
    
    ElseIf i < 153 * 3 And i >= 153 * 2 Then
        For j = 0 To 8
            AF_3(j, counter_k3) = Shuffle_Attributes(j, i)
        Next
        counter_k3 = counter_k3 + 1
    
    ElseIf i < 153 * 4 And i >= 153 * 3 Then
        For j = 0 To 8
            AF_4(j, counter_k4) = Shuffle_Attributes(j, i)
        Next
        counter_k4 = counter_k4 + 1
    
    ElseIf i >= 153 * 4 Then
        For j = 0 To 8
            AF_5(j, counter_k5) = Shuffle_Attributes(j, i)
        Next
        counter_k5 = counter_k5 + 1
    End If
Next

' ================================
' Step 6: Naive Baysian Classifier
' ================================
train_dataset_type_1 = map_into_train_dataset(153, 153, 153, 156, AF_2, AF_3, AF_4, AF_5)
train_dataset_type_2 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_3, AF_4, AF_5)
train_dataset_type_3 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_2, AF_4, AF_5)
train_dataset_type_4 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_2, AF_3, AF_5)
train_dataset_type_5 = map_into_train_dataset(153, 153, 153, 153, AF_1, AF_2, AF_3, AF_4)

fold_1_correct = naive_bayes_classifier(AF_1, 153, train_dataset_type_1, 615)
fold_2_correct = naive_bayes_classifier(AF_2, 153, train_dataset_type_2, 615)
fold_3_correct = naive_bayes_classifier(AF_3, 153, train_dataset_type_3, 615)
fold_4_correct = naive_bayes_classifier(AF_4, 153, train_dataset_type_4, 615)
fold_5_correct = naive_bayes_classifier(AF_5, 156, train_dataset_type_5, 612)

'Checker
'List1.AddItem Str(fold_1_correct) & " " & Str(fold_2_correct) & " " & Str(fold_3_correct) & " " & Str(fold_4_correct) & " " & Str(fold_5_correct)

' ==================================
' Step 7: Calculate Average Accuracy
' ==================================
P2_BAR = (fold_1_correct + fold_2_correct + fold_3_correct + fold_4_correct + fold_5_correct) / 768
List1.AddItem "[Model 2: Equal-frequency Discretization + Naive Bayes Classifier] Average is" & " " & Str(P2_BAR)

' ======================================
' Step 8: Statistical hypothesis testing
' ======================================
List1.AddItem ""
List1.AddItem "----------------------------------------------------------------------------------------------------------------"
List1.AddItem "Perform Statistical hypothesis testing for Model 1 and Model 2"
List1.AddItem "----------------------------------------------------------------------------------------------------------------"
List1.AddItem "Model 1: Equal-width Discretization + Naive Bayes Classifier"
List1.AddItem "Model 2: Equal-Frequency Discretization + Naive Bayes Classifier"
List1.AddItem "H_0: P1_BAR = P2_BAR (Null Hypothesis)"

P_AVG = (P1_BAR + P2_BAR) / 2
Z = (P1_BAR - P2_BAR) / Sqr(2 / 768 * P_AVG * (1 - P_AVG))
List1.AddItem "Z =" & " " & Str(Z)
List1.AddItem "Two-tailed test with alpha = 0.05"
If Z < -1.96 Or Z > 1.96 Then
    List1.AddItem "H_0 is rejected"
Else
    List1.AddItem "H_0 is not rejected"
End If

' ===========================================================================================================

List1.AddItem ""
List1.AddItem "==============================================================="
List1.AddItem "Model 3: K-Nearest Neighbor + Majority Voting"
List1.AddItem "==============================================================="
List1.AddItem ""

' ====================
' Step 1: Read Data
' ====================
in_file = App.Path & "\" & infile.Text
Open in_file For Input As #1
i = 0
Do While Not EOF(1)
    Input #1, feat_x(0), feat_x(1), feat_x(2), feat_x(3), feat_x(4), feat_x(5), feat_x(6), feat_x(7), feat_y
    
    For j = 0 To 7
        Attributes(j, i) = feat_x(j)
    Next j
    
    Attributes(8, i) = feat_y
    i = i + 1
Loop
Close #1

' ======================
' Step 2: Normalize Data
' ======================
For i = 0 To 767
    specific_col_0(i) = Attributes(0, i)
    specific_col_1(i) = Attributes(1, i)
    specific_col_2(i) = Attributes(2, i)
    specific_col_3(i) = Attributes(3, i)
    specific_col_4(i) = Attributes(4, i)
    specific_col_5(i) = Attributes(5, i)
    specific_col_6(i) = Attributes(6, i)
    specific_col_7(i) = Attributes(7, i)
Next

max_attr = find_max(specific_col_0)
min_attr = find_min(specific_col_0)
For i = 0 To 767
    specific_col_0_norm(i) = (specific_col_0(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_1)
min_attr = find_min(specific_col_1)
For i = 0 To 767
    specific_col_1_norm(i) = (specific_col_1(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_2)
min_attr = find_min(specific_col_2)
For i = 0 To 767
    specific_col_2_norm(i) = (specific_col_2(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_3)
min_attr = find_min(specific_col_3)
For i = 0 To 767
    specific_col_3_norm(i) = (specific_col_3(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_4)
min_attr = find_min(specific_col_4)
For i = 0 To 767
    specific_col_4_norm(i) = (specific_col_4(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_5)
min_attr = find_min(specific_col_5)
For i = 0 To 767
    specific_col_5_norm(i) = (specific_col_5(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_6)
min_attr = find_min(specific_col_6)
For i = 0 To 767
    specific_col_6_norm(i) = (specific_col_6(i) - min_attr) / (max_attr - min_attr)
Next

max_attr = find_max(specific_col_7)
min_attr = find_min(specific_col_7)
For i = 0 To 767
    specific_col_7_norm(i) = (specific_col_7(i) - min_attr) / (max_attr - min_attr)
Next

For i = 0 To 767
    Attributes_Norm(0, i) = specific_col_0_norm(i)
    Attributes_Norm(1, i) = specific_col_1_norm(i)
    Attributes_Norm(2, i) = specific_col_2_norm(i)
    Attributes_Norm(3, i) = specific_col_3_norm(i)
    Attributes_Norm(4, i) = specific_col_4_norm(i)
    Attributes_Norm(5, i) = specific_col_5_norm(i)
    Attributes_Norm(6, i) = specific_col_6_norm(i)
    Attributes_Norm(7, i) = specific_col_7_norm(i)
    Attributes_Norm(8, i) = Attributes(8, i)
Next

' ======================
' Step 3: Random Shuffle
' ======================
'For i = 0 To 767
'    random_list(i) = Rnd()
'Next
'arr_ori_idx = sort_col_get_idx(random_list)

'Checker
'For i = 0 To 767
'    List1.AddItem arr_ori_idx(i)
'Next

counter = 0
For ii = 0 To 767
    ori_idx = arr_ori_idx(counter)
    For jj = 0 To 8
        Shuffle_Attributes(jj, counter) = Attributes_Norm(jj, ii)
    Next
    counter = counter + 1
Next

' ==========================
' Step 4: Split into 5 folds
' ==========================
counter_k1 = 0
counter_k2 = 0
counter_k3 = 0
counter_k4 = 0
counter_k5 = 0

For i = 0 To 767
    If i < 153 * 1 Then
        For j = 0 To 8
            AF_1(j, counter_k1) = Shuffle_Attributes(j, i)
        Next
        counter_k1 = counter_k1 + 1
        
    ElseIf i < 153 * 2 And i >= 153 * 1 Then
        For j = 0 To 8
            AF_2(j, counter_k2) = Shuffle_Attributes(j, i)
        Next
        counter_k2 = counter_k2 + 1
    
    ElseIf i < 153 * 3 And i >= 153 * 2 Then
        For j = 0 To 8
            AF_3(j, counter_k3) = Shuffle_Attributes(j, i)
        Next
        counter_k3 = counter_k3 + 1
    
    ElseIf i < 153 * 4 And i >= 153 * 3 Then
        For j = 0 To 8
            AF_4(j, counter_k4) = Shuffle_Attributes(j, i)
        Next
        counter_k4 = counter_k4 + 1
    
    ElseIf i >= 153 * 4 Then
        For j = 0 To 8
            AF_5(j, counter_k5) = Shuffle_Attributes(j, i)
        Next
        counter_k5 = counter_k5 + 1
    End If
Next

' ======================
' Step 5: KNN Classifier
' ======================
train_dataset_type_1 = map_into_train_dataset(153, 153, 153, 156, AF_2, AF_3, AF_4, AF_5)
train_dataset_type_2 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_3, AF_4, AF_5)
train_dataset_type_3 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_2, AF_4, AF_5)
train_dataset_type_4 = map_into_train_dataset(153, 153, 153, 156, AF_1, AF_2, AF_3, AF_5)
train_dataset_type_5 = map_into_train_dataset(153, 153, 153, 153, AF_1, AF_2, AF_3, AF_4)

fold_1_correct = KNN_classifier_type_614(AF_1, 153, train_dataset_type_1, 615, "Majority_Voting")
fold_2_correct = KNN_classifier_type_614(AF_2, 153, train_dataset_type_2, 615, "Majority_Voting")
fold_3_correct = KNN_classifier_type_614(AF_3, 153, train_dataset_type_3, 615, "Majority_Voting")
fold_4_correct = KNN_classifier_type_614(AF_4, 153, train_dataset_type_4, 615, "Majority_Voting")
fold_5_correct = KNN_classifier_type_611(AF_5, 156, train_dataset_type_5, 612, "Majority_Voting")

'List1.AddItem Str(fold_1_correct) & " " & Str(fold_2_correct) & " " & Str(fold_3_correct) & " " & Str(fold_4_correct) & " " & Str(fold_5_correct)
' ==================================
' Step 6: Calculate Average Accuracy
' ==================================
P3_BAR = (fold_1_correct + fold_2_correct + fold_3_correct + fold_4_correct + fold_5_correct) / 768
List1.AddItem "[Model 3: KNN Classifier + Majority Voting] Average is" & " " & Str(P3_BAR)
' ===========================================================================================================

List1.AddItem ""
List1.AddItem "==============================================================="
List1.AddItem "Model 4: K-Nearest Neighbor + Weighted-Distance Voting"
List1.AddItem "==============================================================="
List1.AddItem ""

' ==================================
' No Need to perform Step 1 ~ Step 5
' ==================================

fold_1_correct = KNN_classifier_type_614(AF_1, 153, train_dataset_type_1, 615, "Weighted_Voting")
fold_2_correct = KNN_classifier_type_614(AF_2, 153, train_dataset_type_2, 615, "Weighted_Voting")
fold_3_correct = KNN_classifier_type_614(AF_3, 153, train_dataset_type_3, 615, "Weighted_Voting")
fold_4_correct = KNN_classifier_type_614(AF_4, 153, train_dataset_type_4, 615, "Weighted_Voting")
fold_5_correct = KNN_classifier_type_611(AF_5, 156, train_dataset_type_5, 612, "Weighted_Voting")

'List1.AddItem Str(fold_1_correct) & " " & Str(fold_2_correct) & " " & Str(fold_3_correct) & " " & Str(fold_4_correct) & " " & Str(fold_5_correct)
' ==================================
' Step 6: Calculate Average Accuracy
' ==================================
P4_BAR = (fold_1_correct + fold_2_correct + fold_3_correct + fold_4_correct + fold_5_correct) / 768
List1.AddItem "[Model 4: KNN Classifier + Weighted Voting] Average is" & " " & Str(P4_BAR)

' ======================================
' Step 7: Statistical hypothesis testing
' ======================================
List1.AddItem ""
List1.AddItem "----------------------------------------------------------------------------------------------------------------"
List1.AddItem "Perform Statistical hypothesis testing for Model 3 and Model 4"
List1.AddItem "----------------------------------------------------------------------------------------------------------------"
List1.AddItem "Model 3: K-Nearest Neighbor + Majority Voting"
List1.AddItem "Model 4: K-Nearest Neighbor + Weighted Voting"
List1.AddItem "H_0: P3_BAR = P4_BAR (Null Hypothesis)"

P_AVG = (P3_BAR + P4_BAR) / 2
Z = (P3_BAR - P4_BAR) / Sqr(2 / 768 * P_AVG * (1 - P_AVG))
List1.AddItem "Z =" & " " & Str(Z)
List1.AddItem "Two-tailed test with alpha = 0.05"
If Z < -1.96 Or Z > 1.96 Then
    List1.AddItem "H_0 is rejected"
Else
    List1.AddItem "H_0 is not rejected"
End If

' ====================================================
' Step 8: Statistical Hypothesis Testing For NBC & KNN
' ====================================================
Dim NBC_HIGH As Double
Dim KNN_HIGH As Double
Dim NBC_HIGH_STR As String
Dim KNN_HIGH_STR As String

List1.AddItem ""
List1.AddItem "----------------------------------------------------------------------------------------------------------------"
List1.AddItem "Perform Statistical hypothesis testing for Naive Bayes Classifier (Model 1 & 2) and KNN Classifier (Model 3 & 4)"
List1.AddItem "----------------------------------------------------------------------------------------------------------------"
If P1_BAR > P2_BAR Then
    NBC_HIGH = P1_BAR
    NBC_HIGH_STR = "P1_BAR"
    List1.AddItem "Model 1 performed better than Model 2, so Model 1 is chosen for hypothesis testing."
Else
    NBC_HIGH = P2_BAR
    NBC_HIGH_STR = "P2_BAR"
    List1.AddItem "Model 2 performed better than Model 1, so Model 2 is chosen for hypothesis testing."
End If

If P3_BAR > P4_BAR Then
    KNN_HIGH = P3_BAR
    KNN_HIGH_STR = "P3_BAR"
    List1.AddItem "Model 3 performed better than Model 4, so Model 3 is chosen for hypothesis testing."
Else
    NBC_HIGH = P4_BAR
    KNN_HIGH_STR = "P4_BAR"
    List1.AddItem "Model 4 performed better than Model 3, so Model 4 is chosen for hypothesis testing."
End If

List1.AddItem "H_0:" & "" & NBC_HIGH_STR & "" & "=" & "" & KNN_HIGH_STR & "" & " (Null Hypothesis)"
P_AVG = (NBC_HIGH + KNN_HIGH) / 2
Z = (NBC_HIGH - KNN_HIGH) / Sqr(2 / 768 * P_AVG * (1 - P_AVG))
List1.AddItem "Z =" & " " & Str(Z)
List1.AddItem "Two-tailed test with alpha = 0.05"
If Z < -1.96 Or Z > 1.96 Then
    List1.AddItem "H_0 is rejected"
Else
    List1.AddItem "H_0 is not rejected"
End If

End Sub
Private Function KNN_classifier_type_614(TSTA, test_num, TRNA, train_num, mode)
    Dim ACC_Count As Integer
    Dim test_instance(7) As Double
    Dim diff_list(614) As Double
    Dim diff_0 As Double
    Dim diff_1 As Double
    Dim diff_2 As Double
    Dim diff_3 As Double
    Dim diff_4 As Double
    Dim diff_5 As Double
    Dim diff_6 As Double
    Dim diff_7 As Double
    Dim total_diff As Double
    Dim skip_list(4) As Integer
    Dim min_1_idx As Integer
    Dim min_2_idx As Integer
    Dim min_3_idx As Integer
    Dim min_4_idx As Integer
    Dim min_5_idx As Integer
    Dim vote_1 As Integer
    Dim vote_2 As Integer
    Dim vote_3 As Integer
    Dim vote_4 As Integer
    Dim vote_5 As Integer
    Dim candidate_0 As Integer
    Dim candidate_1 As Integer
    Dim candidate_0_float As Double
    Dim candidate_1_float As Double
    Dim pred_ans As Integer
    
    ACC_Count = 0
    
    For ii = 0 To (test_num - 1)
        test_instance(0) = TSTA(0, ii)
        test_instance(1) = TSTA(1, ii)
        test_instance(2) = TSTA(2, ii)
        test_instance(3) = TSTA(3, ii)
        test_instance(4) = TSTA(4, ii)
        test_instance(5) = TSTA(5, ii)
        test_instance(6) = TSTA(6, ii)
        test_instance(7) = TSTA(7, ii)
        
        For j = 0 To (train_num - 1)
            diff_0 = (test_instance(0) - TRNA(0, j)) * (test_instance(0) - TRNA(0, j))
            diff_1 = (test_instance(1) - TRNA(1, j)) * (test_instance(1) - TRNA(1, j))
            diff_2 = (test_instance(2) - TRNA(2, j)) * (test_instance(2) - TRNA(2, j))
            diff_3 = (test_instance(3) - TRNA(3, j)) * (test_instance(3) - TRNA(3, j))
            diff_4 = (test_instance(4) - TRNA(4, j)) * (test_instance(4) - TRNA(4, j))
            diff_5 = (test_instance(5) - TRNA(5, j)) * (test_instance(5) - TRNA(5, j))
            diff_6 = (test_instance(6) - TRNA(6, j)) * (test_instance(6) - TRNA(6, j))
            diff_7 = (test_instance(7) - TRNA(7, j)) * (test_instance(7) - TRNA(7, j))
            total_diff = diff_0 + diff_1 + diff_2 + diff_3 + diff_4 + diff_5 + diff_6 + diff_7
            diff_list(j) = total_diff
        Next
        
        skip_list(0) = -1
        skip_list(1) = -1
        skip_list(2) = -1
        skip_list(3) = -1
        skip_list(4) = -1
        
        min_1_idx = find_min_pos_type_614(diff_list, skip_list)
        skip_list(0) = min_1_idx
        min_2_idx = find_min_pos_type_614(diff_list, skip_list)
        skip_list(1) = min_2_idx
        min_3_idx = find_min_pos_type_614(diff_list, skip_list)
        skip_list(2) = min_3_idx
        min_4_idx = find_min_pos_type_614(diff_list, skip_list)
        skip_list(3) = min_4_idx
        min_5_idx = find_min_pos_type_614(diff_list, skip_list)
        skip_list(4) = min_5_idx
        'Checker
        'List1.AddItem Str(min_1_idx) & " " & Str(min_2_idx) & " " & Str(min_3_idx) & " " & Str(min_4_idx) & " " & Str(min_5_idx)
    
        If mode = "Majority_Voting" Then
            
            vote_1 = TRNA(8, min_1_idx)
            vote_2 = TRNA(8, min_2_idx)
            vote_3 = TRNA(8, min_3_idx)
            vote_4 = TRNA(8, min_4_idx)
            vote_5 = TRNA(8, min_5_idx)
            
            candidate_0 = 0
            candidate_1 = 0
            
            If vote_1 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_2 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_3 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_4 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_5 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If candidate_0 > candidate_1 Then
                pred_ans = 0
            Else
                pred_ans = 1
            End If
            
            ' Checker
            'List1.AddItem Str(pred_ans) & " " & Str(TSTA(8, i))
            'List1.AddItem Str(TSTA(8, i))
            'List1.AddItem Str(i)
            If pred_ans = TSTA(8, ii) Then
                ACC_Count = ACC_Count + 1
            Else
                ACC_Count = ACC_Count
            End If

        ElseIf mode = "Weighted_Voting" Then
            
            vote_1 = TRNA(8, min_1_idx)
            vote_2 = TRNA(8, min_2_idx)
            vote_3 = TRNA(8, min_3_idx)
            vote_4 = TRNA(8, min_4_idx)
            vote_5 = TRNA(8, min_5_idx)
            
            diff_1 = diff_list(min_1_idx)
            diff_2 = diff_list(min_2_idx)
            diff_3 = diff_list(min_3_idx)
            diff_4 = diff_list(min_4_idx)
            diff_5 = diff_list(min_5_idx)
            
            candidate_0_float = 0
            candidate_1_float = 0
            
            If vote_1 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_1 * diff_1)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_1 * diff_1)
            End If
            
            If vote_2 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_2 * diff_2)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_2 * diff_2)
            End If
            
            If vote_3 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_3 * diff_3)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_3 * diff_3)
            End If
            
            If vote_4 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_4 * diff_4)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_4 * diff_4)
            End If
            
            If vote_5 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_5 * diff_5)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_5 * diff_5)
            End If
            
            'List1.AddItem Str(candidate_0_float) & " " & Str(candidate_1_float)
            
            If candidate_0_float > candidate_1_float Then
                pred_ans = 0
            Else
                pred_ans = 1
            End If
            
            If pred_ans = TSTA(8, ii) Then
                ACC_Count = ACC_Count + 1
            Else
                ACC_Count = ACC_Count
            End If
            
            
        End If
        
    Next
    KNN_classifier_type_614 = ACC_Count
End Function
Private Function KNN_classifier_type_611(TSTA, test_num, TRNA, train_num, mode)
    Dim ACC_Count As Integer
    Dim test_instance(7) As Double
    Dim diff_list(611) As Double
    Dim diff_0 As Double
    Dim diff_1 As Double
    Dim diff_2 As Double
    Dim diff_3 As Double
    Dim diff_4 As Double
    Dim diff_5 As Double
    Dim diff_6 As Double
    Dim diff_7 As Double
    Dim total_diff As Double
    Dim skip_list(4) As Integer
    Dim min_1_idx As Integer
    Dim min_2_idx As Integer
    Dim min_3_idx As Integer
    Dim min_4_idx As Integer
    Dim min_5_idx As Integer
    Dim vote_1 As Integer
    Dim vote_2 As Integer
    Dim vote_3 As Integer
    Dim vote_4 As Integer
    Dim vote_5 As Integer
    Dim candidate_0 As Integer
    Dim candidate_1 As Integer
    Dim candidate_0_float As Double
    Dim candidate_1_float As Double
    Dim pred_ans As Integer
    
    ACC_Count = 0
    
    For ii = 0 To (test_num - 1)
        test_instance(0) = TSTA(0, ii)
        test_instance(1) = TSTA(1, ii)
        test_instance(2) = TSTA(2, ii)
        test_instance(3) = TSTA(3, ii)
        test_instance(4) = TSTA(4, ii)
        test_instance(5) = TSTA(5, ii)
        test_instance(6) = TSTA(6, ii)
        test_instance(7) = TSTA(7, ii)
        
        For j = 0 To (train_num - 1)
            diff_0 = (test_instance(0) - TRNA(0, j)) * (test_instance(0) - TRNA(0, j))
            diff_1 = (test_instance(1) - TRNA(1, j)) * (test_instance(1) - TRNA(1, j))
            diff_2 = (test_instance(2) - TRNA(2, j)) * (test_instance(2) - TRNA(2, j))
            diff_3 = (test_instance(3) - TRNA(3, j)) * (test_instance(3) - TRNA(3, j))
            diff_4 = (test_instance(4) - TRNA(4, j)) * (test_instance(4) - TRNA(4, j))
            diff_5 = (test_instance(5) - TRNA(5, j)) * (test_instance(5) - TRNA(5, j))
            diff_6 = (test_instance(6) - TRNA(6, j)) * (test_instance(6) - TRNA(6, j))
            diff_7 = (test_instance(7) - TRNA(7, j)) * (test_instance(7) - TRNA(7, j))
            total_diff = diff_0 + diff_1 + diff_2 + diff_3 + diff_4 + diff_5 + diff_6 + diff_7
            diff_list(j) = total_diff
        Next
        
        skip_list(0) = -1
        skip_list(1) = -1
        skip_list(2) = -1
        skip_list(3) = -1
        skip_list(4) = -1
        
        min_1_idx = find_min_pos_type_611(diff_list, skip_list)
        skip_list(0) = min_1_idx
        min_2_idx = find_min_pos_type_611(diff_list, skip_list)
        skip_list(1) = min_2_idx
        min_3_idx = find_min_pos_type_611(diff_list, skip_list)
        skip_list(2) = min_3_idx
        min_4_idx = find_min_pos_type_611(diff_list, skip_list)
        skip_list(3) = min_4_idx
        min_5_idx = find_min_pos_type_611(diff_list, skip_list)
        skip_list(4) = min_5_idx
        'Checker
        'List1.AddItem Str(min_1_idx) & " " & Str(min_2_idx) & " " & Str(min_3_idx) & " " & Str(min_4_idx) & " " & Str(min_5_idx)
    
        If mode = "Majority_Voting" Then
            
            vote_1 = TRNA(8, min_1_idx)
            vote_2 = TRNA(8, min_2_idx)
            vote_3 = TRNA(8, min_3_idx)
            vote_4 = TRNA(8, min_4_idx)
            vote_5 = TRNA(8, min_5_idx)
            
            candidate_0 = 0
            candidate_1 = 0
            
            If vote_1 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_2 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_3 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_4 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If vote_5 = 0 Then
                candidate_0 = candidate_0 + 1
            Else
                candidate_1 = candidate_1 + 1
            End If
            
            If candidate_0 > candidate_1 Then
                pred_ans = 0
            Else
                pred_ans = 1
            End If
            
            ' Checker
            'List1.AddItem Str(pred_ans) & " " & Str(TSTA(8, i))
            'List1.AddItem Str(TSTA(8, i))
            'List1.AddItem Str(i)
            If pred_ans = TSTA(8, ii) Then
                ACC_Count = ACC_Count + 1
            Else
                ACC_Count = ACC_Count
            End If

        ElseIf mode = "Weighted_Voting" Then
            
            vote_1 = TRNA(8, min_1_idx)
            vote_2 = TRNA(8, min_2_idx)
            vote_3 = TRNA(8, min_3_idx)
            vote_4 = TRNA(8, min_4_idx)
            vote_5 = TRNA(8, min_5_idx)
            
            diff_1 = diff_list(min_1_idx)
            diff_2 = diff_list(min_2_idx)
            diff_3 = diff_list(min_3_idx)
            diff_4 = diff_list(min_4_idx)
            diff_5 = diff_list(min_5_idx)
            
            candidate_0_float = 0
            candidate_1_float = 0
            
            If vote_1 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_1 * diff_1)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_1 * diff_1)
            End If
            
            If vote_2 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_2 * diff_2)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_2 * diff_2)
            End If
            
            If vote_3 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_3 * diff_3)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_3 * diff_3)
            End If
            
            If vote_4 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_4 * diff_4)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_4 * diff_4)
            End If
            
            If vote_5 = 0 Then
                candidate_0_float = candidate_0_float + 1 / (diff_5 * diff_5)
            Else
                candidate_1_float = candidate_1_float + 1 / (diff_5 * diff_5)
            End If
            
            If candidate_0_float > candidate_1_float Then
                pred_ans = 0
            Else
                pred_ans = 1
            End If
            
            If pred_ans = TSTA(8, ii) Then
                ACC_Count = ACC_Count + 1
            Else
                ACC_Count = ACC_Count
            End If
            
        End If
        
    Next
    KNN_classifier_type_611 = ACC_Count
End Function
Private Function naive_bayes_classifier(TSTA, test_num, TRNA, train_num)
    Dim ACC_Count As Integer
    Dim count_c1 As Double
    Dim count_c2 As Double
    Dim prob_c1 As Double
    Dim prob_c2 As Double
    Dim cond_prob_c1 As Double
    Dim cond_prob_c2 As Double
    Dim target_count_c1 As Double
    Dim target_count_c2 As Double
    Dim score_c1 As Double
    Dim score_c2 As Double
    
    ACC_Count = 0
    
    For i = 0 To (test_num - 1)
    
        count_c1 = 0
        count_c2 = 0
    
        ' 計算 p(c)
        For j = 0 To (train_num - 1)
    
            If TRNA(8, j) = 0 Then
                count_c1 = count_c1 + 1
            End If
    
            If TRNA(8, j) = 1 Then
                count_c2 = count_c2 + 1
            End If
        Next
    
        prob_c1 = count_c1 / train_num
        prob_c2 = count_c2 / train_num
    
        ' 計算 p(x,c)
        cond_prob_c1 = 1
        cond_prob_c2 = 1
        For k = 0 To 7
            target_count_c1 = 0
            target_count_c2 = 0
    
            For ii = 0 To (train_num - 1)
                If (TRNA(8, ii) = 0) And (TSTA(k, i) = TRNA(k, ii)) Then
                    target_count_c1 = target_count_c1 + 1
                ElseIf (TRNA(8, ii) = 1) And (TSTA(k, i) = TRNA(k, ii)) Then
                    target_count_c2 = target_count_c2 + 1
                End If
            Next
    
            cond_prob_c1 = cond_prob_c1 * target_count_c1 / count_c1
            cond_prob_c2 = cond_prob_c2 * target_count_c2 / count_c2
        Next
    
        ' 比較兩個 socre 看誰高
        score_c1 = prob_c1 * cond_prob_c1
        score_c2 = prob_c2 * cond_prob_c2
        If (score_c1 > score_c2) And (TSTA(8, i) = 0) Then
            ACC_Count = ACC_Count + 1
        ElseIf (score_c1 <= score_c2) And (TSTA(8, i) = 1) Then
            ACC_Count = ACC_Count + 1
        End If
    Next
    
    naive_bayes_classifier = ACC_Count
    
End Function
' Helper Function: map_into_train_dataset
Private Function map_into_train_dataset(num_1, num_2, num_3, num_4, AF_1, AF_2, AF_3, AF_4)
    Dim train_dataset(8, 614) As Variant
    Dim counter As Integer
    counter = 0
    For i = 0 To (num_1 - 1)
        For j = 0 To 8
            train_dataset(j, counter) = AF_1(j, i)
        Next
        counter = counter + 1
    Next
    
    For i = 0 To (num_2 - 1)
        For j = 0 To 8
            train_dataset(j, counter) = AF_2(j, i)
        Next
        counter = counter + 1
    Next
    
    For i = 0 To (num_3 - 1)
        For j = 0 To 8
            train_dataset(j, counter) = AF_3(j, i)
        Next
        counter = counter + 1
    Next

    
    For i = 0 To (num_4 - 1)
        For j = 0 To 8
            train_dataset(j, counter) = AF_4(j, i)
        Next
        counter = counter + 1
    Next
    
    map_into_train_dataset = train_dataset
End Function
' Helper Function: sort col get idx
Private Function sort_col_get_idx(arr)
    
    Dim arr_sort(767) As Double
    Dim arr_idx(767) As Variant
    
    ' 先讓 arr_sort 裡面都是-100
    For jj = 0 To 767
        arr_sort(jj) = -100
        arr_idx(jj) = -100
    Next
    
    Dim max_number As Double
    Dim max_number_idx As Integer
    Dim count As Integer
    
    For count = 0 To 767
        
        ' 找到當前 arr 裡面的最大值 以及 他的 index
        max_number = -1000
        max_number_idx = -1000
        
        For kk = 0 To 767
            If arr(kk) > max_number Then
                max_number = arr(kk)
                max_number_idx = kk
            End If
        Next
        
        ' 把它變最小，這樣之後就不會挑到他了
        arr(max_number_idx) = -1000
        
        ' 把當前最大值存到 arr_sort 裡面
        arr_sort(count) = max_number
        arr_idx(count) = max_number_idx
    Next
    sort_col_get_idx = arr_idx
End Function
' Helper Function: find_min
Private Function find_min(col)
    Dim min_num As Double
    Dim ii As Integer
    min_num = 10000
    For ii = 0 To 767
        If col(ii) < min_num Then
            min_num = col(ii)
        End If
    Next
    find_min = min_num
End Function
' Helper Function: find_max
Private Function find_max(col)
    Dim max_num As Double
    Dim ii As Integer
    
    max_num = -999
    For ii = 0 To 767
        If col(ii) > max_num Then
            max_num = col(ii)
        End If
    Next
    find_max = max_num
End Function
' Helper Function: sort col
Private Function sort_col(arr)
    
    Dim arr_sort(767) As Double
    Dim jj As Integer
    
    ' 先讓 arr_sort 裡面都是-100
    For jj = 0 To 767
        arr_sort(jj) = -100
    Next
    
    Dim max_number As Double
    Dim max_number_idx As Integer
    Dim kk As Integer
    
    Dim count As Integer
    
    For count = 0 To 767
        
        ' 找到當前 arr 裡面的最大值 以及 他的 index
        max_number = -1000
        max_number_idx = -1000
        
        For kk = 0 To 767
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
' Helper Function: find min pos
Private Function find_min_pos_type_614(col, skip_list)
    
    Dim var As Collection
    Set var = New Collection
    
    Dim min_num As Double
    Dim min_idx As Double
    min_num = 100000
    min_idx = 100000
    
    Dim toggle As Boolean
    
    For i = 0 To 614
        toggle = False
        For j = 0 To 4
            If i = skip_list(j) Then
                toggle = True
            End If
        Next
        
        If (col(i) < min_num) And (toggle = False) Then
            min_num = col(i)
            min_idx = i
        End If
        
    Next
    
    find_min_pos_type_614 = min_idx
End Function
Private Function find_min_pos_type_611(col, skip_list)
    
    Dim var As Collection
    Set var = New Collection
    
    Dim min_num As Double
    Dim min_idx As Double
    min_num = 100000
    min_idx = 100000
    
    Dim toggle As Boolean
    
    For i = 0 To 611
        toggle = False
        For j = 0 To 4
            If i = skip_list(j) Then
                toggle = True
            End If
        Next
        
        If (col(i) < min_num) And (toggle = False) Then
            min_num = col(i)
            min_idx = i
        End If
        
    Next
    
    find_min_pos_type_611 = min_idx
End Function
