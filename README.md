# JP Morgan Chase Virtual Internship
This virtual internship was under JP Morgan Chase and Co. on Excel Skills by Forage.


### Account Sales Data for Analysis_Task2
* Highlight any cells with formula errors in purple with white text.
* 條件式格式設定-新增規則-樣式（古典）-只格式化包含下列的儲存格-錯誤值
* Highlight any cells with missing values in yellow.
* 條件式格式設定-新增規則-樣式（古典）-只格式化包含下列的儲存格-空格
* Identify accounts that have not been cross-sold with Product 2 by highlighting the appropriate Product 2 cells in orange.
* 條件式格式設定-新增規則-樣式（古典）-醒目提示儲存格規則-等於NO
* Identify accounts that have a 5-year sales CAGR of at least 100% by highlighting the appropriate CAGR cells in green and any accounts with a negative CAGR in red with white text.
* 條件式格式設定-醒目提示儲存格規則-大於等於100％
* 條件式格式設定-醒目提示儲存格規則-小於0％
* Identify accounts in the top 10% of unit sales for 2021 by highlighting the appropriate 2021 unit sales cells in blue.
* 選取2021那一欄-條件式格式設定-前段／後段項目規則-前10％

<img width="823" alt="Account Sales Data for Analysis_Task2" src="https://github.com/user-attachments/assets/6a8844e6-006c-4854-aeb6-17b84c48189c">

### Account Sales Data for Analysis_Task3
You will create two macros and associated buttons:
* A macro to sort the entire spreadsheet by 5 YR CAGR in descending order to see which accounts have the highest overall 5-year sales growth
* 開發人員-錄製巨集-點5YR CAGR下面那一格-排序與篩選-從Z到A-結束錄製-巨集-按鈕

* A macro to sort the entire spreadsheet by 2021 unit sales in descending order to see which accounts have the highest overall unit sales in 2021
* 開發人員-錄製巨集-選取2017到2021下面的資料-排序與篩選-自訂排序-選擇Q欄（2021）-從大到小-結束錄製-巨集-按鈕
儲存檔案的時候，要選擇取用巨集的活頁簿.xlsm的格式

<img width="976" alt="image" src="https://github.com/user-attachments/assets/7c378556-0672-4e20-9f5e-9371e4bcf8c0">

VBA code from Account Sales Data for Analysis_Task3 macros
<img width="1440" alt="image" src="https://github.com/user-attachments/assets/ffa38a75-8851-4179-97af-f453755f28f2">

* A macro to sort the entire spreadsheet by 5 YR CAGR in descending order to see which accounts have the highest overall 5-year sales growth

```
Sub SortBy5YRCAGR()
'
' SortBy5YRCAGR 巨集
'
' 快速鍵: Ctrl+Shift+A
'
    Range("R5").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("R5:R64") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A5:R64")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

```

* A macro to sort the entire spreadsheet by 2021 unit sales in descending order to see which accounts have the highest overall unit sales in 2021

```
Sub SortBy2021UnitSales()
'
' SortBy2021UnitSales 巨集
'
' 快速鍵: Ctrl+Shift+B
'
    Range("M5:Q64").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("Q5:Q64") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("M5:Q64")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

```
