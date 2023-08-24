Attribute VB_Name = "РРЦ_Zyxel"
Public Sub d()
    Dim lLastRow As Long
    'Копируем столбец с GPL
    ActiveSheet.Columns(25).Select '25 - номер столбца
    Selection.Copy
    Selection.Insert
    
    'Вставляем формулу ВПР
    Range("Z4").Select 'Нужно определить первую ячейку
    ActiveCell.FormulaLocal = "=ВПР(F4;'C:\Users\user\Desktop\!_GPLи для компани\[Vendor_2023-03.xlsx]Sheet'!$A:$J;10;ЛОЖЬ)/$AE$1" 'Для каждого прайса своя формула
    lLastRow = Cells(Rows.Count, 25).End(xlUp).Row 'Определяем последнюю строку диапазона
    Range("Z4:Z" & lLastRow).FillDown 'Растягиваем формулу до конца столбца
    Application.Sleep 5000 'сон на 5 сек

    Dim rngColor As Range
    Dim rngCellVisible As Range
    Dim lastRow As Long
    'Фильтруем новый столбец GPL по значению #Н/Д
    ActiveSheet.Range("A1").AutoFilter Field:=26, Criteria1:="#Н/Д"
    'Фильтруем старый столбец GPL по значению не равному #Н/Д
    ActiveSheet.Range("A1").AutoFilter Field:=25, Criteria1:="<>#N/A"
    
    'Переношу старые цены в новый столбец GPL
    lastRow = Cells(Rows.Count, 25).End(xlUp).Row
    For Each rngCellVisible In Range("Y4:Y" & lastRow).SpecialCells(xlCellTypeVisible)
        rngCellVisible.Offset(0, 1).Value = rngCellVisible.Value
        rngCellVisible.Offset(0, 1).Interior.Color = 192
    Next rngCellVisible

    'Убираю фильтры с листа
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Names("_FilterDatabase").Delete
    'Выделяю новый столбец GPL от первой ячейки с ВПР и до последней заполненной ячейки и копирую
    Range("Z4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    'Вставляю только значения, без формулы ВПР
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Удаляю старый столбец GPL
    Columns(25).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    'Меняю дату обновления GPL
    Range("Y3").Select
    ActiveCell.Value = Date
End Sub

