# 此為 VB 學習筆記
* 內容大量參考 [電腦程式設計教學網頁 劉陳祥老師](http://web.tnu.edu.tw/me/study/moodle/tutor/vb6/index.html)

# 第 01 單元：VB 的基本概念
* 點選 檢視 > 即時運算視窗 能夠快速實驗程式片段

# 第 02 單元：VB 的資料處理
1. 資料運算
    * 串接運算
    * 比較運算
    * 邏輯運算
    * 綜合運算
2. 資料類型：宣告、不同型別資料的轉換與運算
    * 數值型資料：
        * 整數(Integer)：%
            * `Dim A%` '將變數 A 宣告成整數型別
        * 長整數(Long)：&
        * 倍精準度型(Double)
        * 單精準度型(Single)：!
    * 字串型資料(String)：$
        * `Dim S$` '將變數 S 宣告成整型別
        * `Dim S As String` '將變數 S 宣告成字串型別
        * 固定長度字串
            * `Dim S As String*80` '指定字串長度=80
    * 日期時間型資料(Date)：無
        * `Dim X As Date` '將變數 X 宣告成日期時間型別
    * 布林(Boolean)：無
    * 不定型
        * `Dim V` '省略 As, 則變數 V 被宣告成不定型變數
    * 常數符號的定義
        * `Const pi = 3.141593`
        * `Const ver = "6.0中文版"`
        * `Const noon = #12:00:00#`
    * 資料型別轉換
        * 使用型別轉換函數：
            * 把字串變成整數：`I% = Val("123")`，其實也可以這樣寫: `I% = "123"`
            * 把整數變成字串：`S$ = Str(123)`，其實也可以這樣寫：`S$ = 123`
            * 以下為錯誤示範：`I% = "123A"`

# 第 03 單元：VB 的基本語法
### 設值語法：
在 VB 中設值語法有兩種：
1. 對一般變數的設值：
    * 變數 = 敘述式
    * `Dim I As Integer`
    * `I = 60*20`
2. 對指定變數的設值：
    * Set 變數 = 敘述式
    * `Dim Ex As DataBase`
    * `Set Ex = OpenDataBase("File.mdb")`

### If 語法：
If 語法用於判斷條件，根據判斷的結果，執行不同的敘述。
1. 格式 I：
```
If 敘述式 Then
...
Else
...
End If
```
2. 格式 II：
```
If 敘述式 Then 語法 ...
```
3. 舉例：
```
If Password="ABC1234" Then
    OK = True
Else
    OK = False
End if
```

### Select Case 語法：
Selct Case 語法用於對某一敘述式的值進行多種判斷處理。
```
Select4 Case 敘述式
Case 值1:
...
Case 值2:
...
Case Else
...
End Select
```

### For 迴圈語法：
For ... Next 用於指定次數的迴圈，格式有兩種：
1. 一般的數字變數：
```
For var = start To end[Step step]
...
Next var
```

2. 指定作用對象：
```
For Each obj In objs
...
Next obj
```
其中：obj 是對象變數，而 objs 是集合變數。

### Do 迴圈語法：
根據條件成立與否來決定是否繼續執行 Do 迴圈，Do 迴圈有兩種：
1. 先判斷後執行：
```
Do While|Until 條件
...
Loop
```

2. 先執行後判斷：
```
Do
...
Loop While|Until 條件
```
Note：While 當條件為 True 時繼續執行迴圈，Until 當條件為 True 時退出。

### With 語法：
當我們經常使用某一對象的屬性、方法時，就可以使用 With 語法，With 語法可以使程式碼更簡潔，還可以提高執行速度。

1. 格式：
```
With 對象變數
...
End With
```

2. 舉例：
```
With Text1
    .SelStart=0
    .SeiLength=Len(.Text)
    .SetFocus
End With
```

3. 相當於：
```
Text1.SelStart=0
Text1.SeiLength=Len(Text1.Text)
Text1.SetFocus
``