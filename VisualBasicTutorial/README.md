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
* 設值語法：
    1. 對一般變數的設值：
        * 變數 = 敘述式
        * `Dim I As Integer`
        * `I = 60*20`
    2. 對指定變數的設值：
        * Set 變數 = 敘述式
        * `Dim Ex As DataBase`
        * `Set Ex = OpenDataBase("File.mdb")`
* If 語法：
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