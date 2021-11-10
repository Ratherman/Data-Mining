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
If 條件式 Then
    指令敘述 1......
Else
...
End If
```
2. 格式 II：
```
If 條件式 Then 指令敘述
```
3. 舉例：
```
If Password="ABC1234" Then
    OK = True
Else
    OK = False
End if
```
4. 格式 III：
```
If 條件式1 Then
    指令敘述A
ElseIf 條件式2 Then
    指令敘述B
ElseIf 條件式3 Then
    指令敘述C
Else
    指令敘述D
End If
```

### Select Case 語法：
Selct Case 語法用於對某一敘述式的值進行多種判斷處理。
```
Select4 Case 條件運算式
Case 測試結果1:
    指令敘述1
Case 測試結果2:
    指令敘述2
Case 測試結果3:
    指令敘述3
...
Case 測試結果N:
    指令敘述N
Case Else
    指令敘述N+1
End Select
```

### For 迴圈語法：
For ... Next 用於指定次數的迴圈，格式有兩種：
1. 一般的數字變數：
```
For 變數 = 初值 To 終值[Step 增值]
    [程式敘述區段]
    [判斷式 ... Exit For]
Next 變數
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
    [程式敘述區段]
    [Exit Do]
Loop
```

2. 先執行後判斷：
```
Do
    [程式敘述區段]
    [Exit Do]
Loop While|Until 條件
```

3. 還有
```
Do Until
    [程式敘述區段]
    [Exit Do]
Loop
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
    .SelLength=Len(.Text)
    .SetFocus
End With
```

3. 相當於：
```
Text1.SelStart=0
Text1.SelLength=Len(Text1.Text)
Text1.SetFocus
```

# 第 04 單元：VB 的程式設計
Visual Basic 撰寫程式的步驟：
1. 選擇【File】>【New Project】開啟新專案檔
2. 從【ToolBasr】中選擇所需的控制項放置到 Form 中的適當位置，並調整其大小。
3. 透過【Properties】視窗設定每一個控制項的相關屬性。
4. 開啟程式碼編寫視窗，轉寫變數宣告、副程式及相關事件程序。
5. 選擇【Run】>【Start】功能(或按 F5)，實際測試程式。
6. 測試正程後，選擇【File】>【Make EXE File】功能，產生執行檔。

# 第 05 單元：VB的物件應用
1. 標籤(Label)
    * Label 控制項元件主要用在視窗中顯示提示訊息，常與 Text 控制項元件一起使用。
    * 屬性：Caption, Font, ForeColor, Enabled, Visible, Top/Left/Width/Height
    * 事件：Click 當點選或按下快捷鍵時發生，常用於和她一起使用的 Text 控制項元件獲得輸入焦點。
```
Private Sub Label1_Click()
Text1.SetFocus
End Sub
```

2. 文字框(Text)
    * Text 控制項元件接受使用者的輸入的字串數據。
    * 屬性：Text, SelStart, SelLength, SelText, MultiLine, ScrollBars, Password, SetFocus, KeyAscii, Change, LostFocus, GotFocus
    * 程式碼：
將Text1控制項接收使用者輸入的數據通過Label1顯示出來。
```
Label1.Caption = Text1.Text
```
GotFocus 通常我們在 Text 控制項元件獲得輸入焦點時全選他的內容，方便使用者直接修改數據。
```
Private Sub Text1_GotFocus()
Text1.SelStart=0
Text1.SelLength=Len(Text1.Text)
End Sub
```

3. 命令按鈕(CommandButton)
    * CommanButton 控制項元件接受使用者的命令。
    * 屬性：Caption (表示按鈕所顯示的內容)
    * 事件：Click (當點選或按下快捷鍵時發生)

4. 核取框(CheckBox)
    * CheckBox 控制項元件檢查某個選項是否被選中。
    * 屬性：Caption, Value
    * 事件：Click (當點選或按下快捷鍵時發生)

5. 選項按鈕(OptionButton)
    * OptionButton 控制項元件檢查一個選項是否被選中，與 checkBox 的區別是 checkBox 是多選多項，而 OptionButton 是多選一項。
    * 屬性：Caption, Value
    * 事件：Click (當點選或按下快捷鍵時發生)

6. 框架(Frame)
    * Frame 控制項元件主要用於 OptionButton 控制項元件分組。
    * 屬性：Caption (表示分組所提示的內容)

7. 清單(ListBox)
    * ListBox 控制項元件是用於在一組列表中選擇其中的一項或多項
    * 屬性：Text, ListCount, ListIndex, List(i), MultiSelect, Selected(i), SelCount, Sort
    * 方法：
        * AddItem 向列表框增加一項數據。`ListX.AddItem(Item As String)`
        * RemoveItem 刪除第 i 項。`ListX.RemoveItem(i As Integer)`
    * 事件：Click (當點選列表框中的一項數據時發生)

8. 下拉式清單(CoomboBox)
    * ComboBox 控制項元件與 ListBox 基本相同，他的優點在佔用的面積小，除了可以在選項中選擇外還可以輸入其他數據。他的缺點是不能多選擇。
    * 屬性：Text (存放從選項中選擇的數據或使用者輸入的數據)。

9. 影像框(Image)
    * Image 控制元件用用於顯示一張圖片。
    * 屬性： Picture (存放圖片的數據)、Stretch (顯示圖片的方式)
        * 通常使用 `LoadPicture` 函數讀入一張圖片。
        * 舉例：`ImageX.Picture = LoadPicture("C:1.bmp")`
        * Note: `LoadPicture` 支持 Bmp, Jpg, Gif 等多種格式圖片文件。

10. 計時器(Timer)
    * Timer 控制項元件以固定間隔時間觸發他的 Timer 事件
    * 屬性：Enabled (表示是否啟動計時器)、Interval (表示觸發 Timer 事件的間隔時間，以毫秒為單位)
    * 事件：Timer 當計時器時到間隔時間時發生。

11. 磁碟機清單(DriveListBox)
    * DriveListBox 控制項元件提供一個驅動器列表。
    * 屬性：Drive (表示當前選擇的驅動器)
    * 事件：Change (當驅動器選擇發生變化時觸發)

12. 目錄清單(DirListBox)
    * DirListBox 控制項元件提供一個目錄列表
    * 屬性：Path (表示當前目錄的路徑)
    * 事件：Change (當目錄選擇發生變化時發生)

13. 檔案清單(FileListBox)
    * FileListBox 控制項元件提供一個文件列表
    * 屬性：Path (表示當前文件列表所在的路徑)、Filename (表示選擇的文件名，不含路徑)、Pattern (決定甚麼樣的文件)
    * 事件：Click (當點選列表框的一項數據時發生)

14. 通用對話框(CommandDialog)
CommandDialog 控制項元件包含了 Windows 操作系統提供的 6 種公用對話框：
    * Open 對話框 和 Save 對話框
    * Color 對話框
    * Font 對話框
    * Printer 對話框
    * Font 對話框

