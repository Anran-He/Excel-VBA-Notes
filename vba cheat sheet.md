# VBA Cheat Sheet
1. copy values
    ```vb
    ws.Range("A1").value = ws.Range("A1").value
    ```
2. get last row number
    ```vb
    lrow = ws.Range("A1").End(xlDown).Row
    ```
3. delete rows
    ```vb
    ws.Rows(2).Delete
    ws.Rows("1:3").Delete
    ```
4. insert formula
    ```VB
    ws.Range("D4").Formula = "=B3*10"
    ws.Range("D4").FormulaR1C1 = "=R3C2*10"
    'relative position
    ws.Range("D4").FormulaR1C1 = "=R[-1]C[-2]*10" 
    ```
5. auto fill
    ```vb
    ws.Range("D4").AutoFill Destination:=ws.Range("D4:D" & lrow)
    ```
6. set workbook
    ```vb
    Set wb = Workbooks.Open(Filename:=ThisWorkbook)
    'read only mode
    Set wb = Workbooks.Open(Filename:=ThisWorkbook, ReadyOnly:=True)
    ```
7. set worksheet
    ```vb
    Set ws = wb.Worksheets(1)
    ```
8. filter
    ```vb
    'don't use any argument, simply apply or remove the filter icons
    ws.Range("A1").AutoFilter
    'filter data based on a text condition
    'filter the 2nd column
    ws.Range("A1").AutoFilter Field:=2, Criteria1:="Printer"
    'filter data on multiple condition
    ws.Range("A1").AutoFilter Field:=2, Criteria1:="Printer", Operator:=xlOr,   Criteria2:="Projector"
    ws.Range("A1").AutoFilter Field:=4, Criteria1:=">10", Operator:=xlAnd, Criteria2:="<20"
    'multiple criteria with different columns
    With ws.Range("A1")
    .AutoFilter Field:=2, Criteria1:="Printer"
    .AutoFilter Field:=3, Criteria1:="Mark"
    'filter top 10 records
    'operator is always "xlTop10Items", if want to get top 5 records, just change criteria from 10 to 5
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="10",
    Operator:=xlTop10Items
    'filter bottom 10 records
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="10",
    Operator:=xlBottom10Items
    'filter top 10 percent records
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="10",
    Operator:=xlTop10Percent
    'copy filtered rows into a new sheet
    Dim rng As Range
    Dim ws As Worksheet
    Set rng = Worksheets("Sheet1").AutoFilter.Range
    Set ws = Worksheets.Add
    rng.Copy Range("A1")
    'turn off auto filter
    ws.AutoFilterMode = False
    'show all data
    ws.ShowAllData
    ```


9. select visible cells (after filter)
    ```vb
    Dim rng As Range
    Set rng = ws.Range("A1:A10").SpecialCells(xlCellTypeVisible).Cells
    rng.Select
    ```

10. set current directory
    ```vb
    ChDir "C:\instructions"
    ```

11. add a new sheet
    ```vb
    Sheets.Add After:=ActiveSheet
    ```

12. add a new sheet and rename
    ```vb
    Sheets.Add(After:=ws).Name="SheetName"
    ```

12. convert text to number
    ```vb
    ws.Range("B:B").NumberFormat = "General"
    ws.Range("B:B").Value = ws.Range("B:B").Value
    ```

13. Paste - Transpose
    ```vb
    ws.Range("A:A").Copy
    ws.Range("B1").PasteSpecial Transpose:=True
    ```

14. Handle error
    ```vb
    On Error Resume Next
    ```