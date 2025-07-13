# my-first-repo
院區陽性vba
Sub 合併微生物報告_插入空白欄修正版()
    Dim fd As FileDialog
    Dim selectedFiles() As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim destWB As Workbook
    Dim destRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim bacteriaCol As Long, mdrCol As Long, mcimCol As Long
    Dim i As Long, idx As Long

    bacteriaCol = 0 ' 預設為0，方便後續判斷

    Set destWB = Workbooks.Add
    Set wsDest = destWB.Sheets(1)
    destRow = 1

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "請選擇要合併的微生物報告 Excel 檔案"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        .AllowMultiSelect = True

        If .Show = -1 Then
            ReDim selectedFiles(1 To .SelectedItems.Count)
            For idx = 1 To .SelectedItems.Count
                selectedFiles(idx) = .SelectedItems(idx)
            Next idx
        Else
            MsgBox "未選擇檔案，已取消操作。", vbExclamation
            Exit Sub
        End If
    End With

    ' 合併資料
    For idx = LBound(selectedFiles) To UBound(selectedFiles)
        Set wbSrc = Workbooks.Open(selectedFiles(idx), ReadOnly:=True)
        Set wsSrc = wbSrc.Sheets(1)

        lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

        If destRow = 1 Then
            ' 複製標題
            wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy wsDest.Cells(destRow, 1)
            destRow = destRow + 1
        End If

        ' 複製資料（不含標題列）
        wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDest.Cells(destRow, 1)
        destRow = destRow + (lastRow - 1)

        wbSrc.Close False
    Next idx

    ' 找到 Bacteria 欄
    For i = 1 To wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
        If Trim(wsDest.Cells(1, i).Value) = "Bacteria" Then
            bacteriaCol = i
            Exit For
        End If
    Next i

    If bacteriaCol = 0 Then
        MsgBox "找不到 [Bacteria] 欄位", vbCritical
        Exit Sub
    End If

    ' 插入兩欄在 Bacteria 右側
    wsDest.Columns(bacteriaCol + 1).Insert Shift:=xlToRight
    wsDest.Columns(bacteriaCol + 1).Insert Shift:=xlToRight

    wsDest.Cells(1, bacteriaCol + 1).Value = "MDR"
    wsDest.Cells(1, bacteriaCol + 2).Value = "mCIM"

    mdrCol = bacteriaCol + 1
    mcimCol = bacteriaCol + 2

    ' 從下往上刪除與標記資料
    For i = wsDest.Cells(wsDest.Rows.Count, bacteriaCol).End(xlUp).Row To 2 Step -1
        Dim cellVal As String
        cellVal = Trim(wsDest.Cells(i, bacteriaCol).Value)

        If LCase(Left(cellVal, 11)) = "results are" Then
            wsDest.Rows(i).Delete
        Else
            If InStr(cellVal, "MDR") > 0 Then
                wsDest.Cells(i, mdrCol).Value = "MDR"
            End If
            If InStr(cellVal, "mCIM(+)") > 0 Then
                wsDest.Cells(i, mcimCol).Value = "mCIM(+)"
            End If
        End If
    Next i

    MsgBox "合併與處理完成！", vbInformation
End Sub
