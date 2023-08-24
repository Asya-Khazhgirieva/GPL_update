Attribute VB_Name = "���_Zyxel"
Public Sub d()
    Dim lLastRow As Long
    '�������� ������� � GPL
    ActiveSheet.Columns(25).Select '25 - ����� �������
    Selection.Copy
    Selection.Insert
    
    '��������� ������� ���
    Range("Z4").Select '����� ���������� ������ ������
    ActiveCell.FormulaLocal = "=���(F4;'C:\Users\user\Desktop\!_GPL� ��� �������\[Vendor_2023-03.xlsx]Sheet'!$A:$J;10;����)/$AE$1" '��� ������� ������ ���� �������
    lLastRow = Cells(Rows.Count, 25).End(xlUp).Row '���������� ��������� ������ ���������
    Range("Z4:Z" & lLastRow).FillDown '����������� ������� �� ����� �������
    Application.Sleep 5000 '��� �� 5 ���

    Dim rngColor As Range
    Dim rngCellVisible As Range
    Dim lastRow As Long
    '��������� ����� ������� GPL �� �������� #�/�
    ActiveSheet.Range("A1").AutoFilter Field:=26, Criteria1:="#�/�"
    '��������� ������ ������� GPL �� �������� �� ������� #�/�
    ActiveSheet.Range("A1").AutoFilter Field:=25, Criteria1:="<>#N/A"
    
    '�������� ������ ���� � ����� ������� GPL
    lastRow = Cells(Rows.Count, 25).End(xlUp).Row
    For Each rngCellVisible In Range("Y4:Y" & lastRow).SpecialCells(xlCellTypeVisible)
        rngCellVisible.Offset(0, 1).Value = rngCellVisible.Value
        rngCellVisible.Offset(0, 1).Interior.Color = 192
    Next rngCellVisible

    '������ ������� � �����
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Names("_FilterDatabase").Delete
    '������� ����� ������� GPL �� ������ ������ � ��� � �� ��������� ����������� ������ � �������
    Range("Z4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    '�������� ������ ��������, ��� ������� ���
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '������ ������ ������� GPL
    Columns(25).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    '����� ���� ���������� GPL
    Range("Y3").Select
    ActiveCell.Value = Date
End Sub

