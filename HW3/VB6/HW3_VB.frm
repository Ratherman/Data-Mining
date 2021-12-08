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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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


' Ū��Ƭ���
Dim in_file As String
Dim feat_y As String
Dim feat_x(9) As Double

' �D�n�x�s��T���G���x�}
Dim Attributes(9, 214) As Variant
Dim attributes_EWD(9, 214) As Variant

' �Φb Equal Width Discretization
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

' �M�� List 1 �����e
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
    
        If j = 0 Or j = 1 Or j = 4 Then
            Attributes(j, i) = feat_x(j + 1) * 100 ' scale �Ӥp�A���n discretization
        Else
            Attributes(j, i) = feat_x(j + 1) * 1000
        End If
        
        'Attributes(j, i) = feat_x(j + 1) * 1000 ' scale �Ӥp�A���n discretization
    Next j
    
    Attributes(9, i) = feat_y
    ' �����C�L�X�Ӭݬ�
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
    
    For k = 0 To 213
        specific_col(k) = Attributes(i, k)
    Next
    
    If i <> 9 Then
        number_of_intervals = 10
        max_num_in_col = find_max(specific_col)
        min_num_in_col = find_min(specific_col)
        List1.AddItem Str(max_num_in_col)
        List1.AddItem Str(min_num_in_col)
        
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
        
        If i = 0 Or i = 1 Or i = 4 Then
           
            sp_str_1 = split_point_1 / 100
            sp_str_2 = split_point_2 / 100
            sp_str_3 = split_point_3 / 100
            sp_str_4 = split_point_4 / 100
            sp_str_5 = split_point_5 / 100
            sp_str_6 = split_point_6 / 100
            sp_str_7 = split_point_7 / 100
            sp_str_8 = split_point_8 / 100
            sp_str_9 = split_point_9 / 100
        
            'List1.AddItem "��" & "" & Str(i) & "" & "�� attribute ���Q�Ӱ϶��Ѥj�ܤp�Ƨ�"
            'List1.AddItem "  �Ĥ@�϶�: [" & "" & Str(sp_str_1) & "" & ", " & "" & "Max num]"
            'List1.AddItem "  �ĤG�϶�: [" & "" & Str(sp_str_2) & "" & ", " & "" & Str(sp_str_1) & "" & ")"
            'List1.AddItem "  �ĤT�϶�: [" & "" & Str(sp_str_3) & "" & ", " & "" & Str(sp_str_2) & "" & ")"
            'List1.AddItem "  �ĥ|�϶�: [" & "" & Str(sp_str_4) & "" & ", " & "" & Str(sp_str_3) & "" & ")"
            'List1.AddItem "  �Ĥ��϶�: [" & "" & Str(sp_str_5) & "" & ", " & "" & Str(sp_str_4) & "" & ")"
            'List1.AddItem "  �Ĥ��϶�: [" & "" & Str(sp_str_6) & "" & ", " & "" & Str(sp_str_5) & "" & ")"
            'List1.AddItem "  �ĤC�϶�: [" & "" & Str(sp_str_7) & "" & ", " & "" & Str(sp_str_6) & "" & ")"
            'List1.AddItem "  �ĤK�϶�: [" & "" & Str(sp_str_8) & "" & ", " & "" & Str(sp_str_7) & "" & ")"
            'List1.AddItem "  �ĤE�϶�: [" & "" & Str(sp_str_9) & "" & ", " & "" & Str(sp_str_8) & "" & ")"
            'List1.AddItem "  �ĤQ�϶�: [" & "" & "Min num" & "" & ", " & "" & Str(sp_str_9) & "" & ")"
            'List1.AddItem ""

        
        Else
            sp_str_1 = split_point_1 / 1000
            sp_str_2 = split_point_2 / 1000
            sp_str_3 = split_point_3 / 1000
            sp_str_4 = split_point_4 / 1000
            sp_str_5 = split_point_5 / 1000
            sp_str_6 = split_point_6 / 1000
            sp_str_7 = split_point_7 / 1000
            sp_str_8 = split_point_8 / 1000
            sp_str_9 = split_point_9 / 1000
            
            'List1.AddItem "��" & "" & Str(i) & "" & "�� attribute ���Q�Ӱ϶��Ѥj�ܤp�Ƨ�"
            'List1.AddItem "  �Ĥ@�϶�: [" & "" & Str(sp_str_1) & "" & ", " & "" & "Max num]"
            'List1.AddItem "  �ĤG�϶�: [" & "" & Str(sp_str_2) & "" & ", " & "" & Str(sp_str_1) & "" & ")"
            'List1.AddItem "  �ĤT�϶�: [" & "" & Str(sp_str_3) & "" & ", " & "" & Str(sp_str_2) & "" & ")"
            'List1.AddItem "  �ĥ|�϶�: [" & "" & Str(sp_str_4) & "" & ", " & "" & Str(sp_str_3) & "" & ")"
            'List1.AddItem "  �Ĥ��϶�: [" & "" & Str(sp_str_5) & "" & ", " & "" & Str(sp_str_4) & "" & ")"
            'List1.AddItem "  �Ĥ��϶�: [" & "" & Str(sp_str_6) & "" & ", " & "" & Str(sp_str_5) & "" & ")"
            'List1.AddItem "  �ĤC�϶�: [" & "" & Str(sp_str_7) & "" & ", " & "" & Str(sp_str_6) & "" & ")"
            'List1.AddItem "  �ĤK�϶�: [" & "" & Str(sp_str_8) & "" & ", " & "" & Str(sp_str_7) & "" & ")"
            'List1.AddItem "  �ĤE�϶�: [" & "" & Str(sp_str_9) & "" & ", " & "" & Str(sp_str_8) & "" & ")"
            'List1.AddItem "  �ĤQ�϶�: [" & "" & "Min num" & "" & ", " & "" & Str(sp_str_9) & "" & ")"
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
            attributes_EWD(i, j) = Attributes(i, j)
        Next
    End If
    
Next
List1.AddItem "���@�U�O�o�n���}"

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
'    ' �@�� accuracy �|������ �@�� experiment �� col�A�b�}�l�e����M��
'    For i = 0 To 8
'        accuracy_list(i) = -1
'        exp_attr_list(i) = -1
'    Next
'
'    ' �i��P�� attr �ƥت��j��
'    count_same_vars = 0
'
'    For j = 0 To 8
'
'        If possible_subset(j) = True Then
'
'            ' �C���j���L�ۦP level �� subset �ɡA�ݭn�M�� chosen_subset
'            ' �b�o��@�֥[�J���窺 attr num
'            For i = 0 To 8
'
'                If i = j Then
'                    chosen_subset(i) = i
'                    exp_attr_list(count_same_vars) = i
'                Else
'                    chosen_subset(i) = -1 ' zero subset ������
'                End If
'            Next
'
'            '�⨺�� �̴Ϊ� attr �[�i��
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
'    ' ��� accuracy�@�̤j�� idx�A�óz�L exp_attr_list ��� attr_num �M goodness
'    idx = argmax(accuracy_list)
'    current_best_attr = exp_attr_list(idx)
'    current_best_accuracy = accuracy_list(idx)
'
'    '�]�w�������
'    ' (1) current_best_accuracy �� best_accuracy �C
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

Dim col_boolean(9) As Boolean

' ����l�� col_name�A���j�a���O False
For i = 0 To 8
    col_boolean(i) = False
Next

' �]�w�Q�n���窺��H�A���]�O 3 �M 0
col_boolean(0) = True
col_boolean(3) = True

acc = selective_naive_bayes(attributes_EWD, col_boolean)

' Test unique
'Dim test_col(213) As Double
'For i = 0 To 212
'    test_col(i) = attributes_EWD(9, i)
'Next
'Dim n As Integer
'n = unique_type_count(test_col)

End Sub
' Core Function: Perform SNB Algorithm with Laplace Estimate
Private Function selective_naive_bayes(Attributes, col_boolean):
    
    ' �ǳƦn Y_col
    Dim Y_col(214) As Double
    For i = 0 To 213
        Y_col(i) = Attributes(9, i)
    Next
    
    ' �ǳƦn Datum�A����L��l�Ʀ� -999
    Dim Datum(214, 9) As Double
    For i = 0 To 213
        For j = 0 To 8
            Datum(i, j) = -999
        Next
    Next
    
    ' �M���ݭn�ϥΪ����(�C�@�����O���@�˪�column)�˶i Datum �̭�
    For i = 0 To 8
        If col_boolean(i) = True Then
            For j = 0 To 213
                Datum(j, i) = Attributes(i, j)
            Next
        End If
    Next
    
    ' �o�� loop �O�b�Ϥ� testing data �M training data �Ϊ�
    Dim ACC_Count As Integer
    ACC_Count = 0
    

    
    For i = 0 To 214
        ' �Ϥ� training data �M testing data
        
        ' ���X testing data
        Dim X_test(1, 9) As Double
        Dim Y_test As Double
        
        For j = 0 To 8
            X_test(0, j) = Datum(i, j)
            Y_test = Y_col(i)
        Next
        
        ' ���X training data
        Dim X_train(213, 9) As Double
        Dim Y_train(213) As Double
        
        For j = 0 To 8
            
            Dim train_counter As Integer
            train_counter = 0
            
            If col_boolean(j) = True Then
                For k = 0 To 214
                    If k = i Then ' ��ܳo����ƬO testing data ��
                        train_counter = train_counter
                    Else
                        X_train(train_counter, j) = Datum(k, j)
                        Y_train(train_counter) = Y_col(k)
                        train_counter = train_counter + 1
                    End If
                Next
            End If
        Next
        
        ' �b i-th run �̭�
        ' �ؼЬO���� X_test �M��� X_train �̭��� Naive Bayes with Lapace Estimate �� Score
        ' �M��A�� Y_test ���
        ' �� 6 �� class type: 1,2,3,5,6,7
        
        ' (1) �p�� p(c)
        Dim count_c1 As Double
        Dim count_c2 As Double
        Dim count_c3 As Double
        Dim count_c5 As Double
        Dim count_c6 As Double
        Dim count_c7 As Double
        
        count_c1 = 0
        count_c2 = 0
        count_c3 = 0
        count_c5 = 0
        count_c6 = 0
        count_c7 = 0
        
        For j = 0 To 213
            
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
        
        Dim prob_c1 As Double
        Dim prob_c2 As Double
        Dim prob_c3 As Double
        Dim prob_c5 As Double
        Dim prob_c6 As Double
        Dim prob_c7 As Double
        
        prob_c1 = count_c1 / 213
        prob_c2 = count_c2 / 213
        prob_c3 = count_c3 / 213
        prob_c5 = count_c5 / 213
        prob_c6 = count_c6 / 213
        prob_c7 = count_c7 / 213
        
    Next

    
    selective_naive_bayes = 1
End Function

' Helper Function: count number of unique types in a column
' Laplace Estimate �ݭn�o�Ӽƭ�(�Φb����)
' �ڥi�H�γo�ؤ覡�h��X unique ���ƶq�O�]���ڤw�g����Ҧ����i��ȴ��� 0 ~ 9 �䤤�@�ӼƦr�F
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
    
    ' �o�̥u���� 213 ����ƬO�]���ڦ��� dataset ������
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
    List1.AddItem Str(num_unique)
    unique_type_count = num_unique
End Function

' Helper Function: sort col
Private Function sort_col(arr)
    
    Dim arr_sort(213) As Double
    Dim jj As Integer
    
    ' ���� arr_sort �̭����O-100
    For jj = 0 To 213
        arr_sort(jj) = -100
    Next
    
    Dim max_number As Double
    Dim max_number_idx As Integer
    Dim kk As Integer
    
    Dim count As Integer
    
    For count = 0 To 213
        
        ' ����e arr �̭����̤j�� �H�� �L�� index
        max_number = -1000
        max_number_idx = -1000
        
        For kk = 0 To 213
            If arr(kk) > max_number Then
                max_number = arr(kk)
                max_number_idx = kk
            End If
        Next
        
        ' �⥦�̤ܳp�A�o�ˤ���N���|�D��L�F
        arr(max_number_idx) = -1000
        
        ' ���e�̤j�Ȧs�� arr_sort �̭�
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

