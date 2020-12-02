

'替儲存格設定顯示格式
Sub Sample8()
    Range("A1").NumberFormatLocal = "$#,###"
End Sub

'替儲存格設定格狀框線
Sub Sample17()
    Range("A2:C4").Borders.LineStyle = xlContinuous
End Sub

'替儲存格設定外框線
Sub Sample18()
    Range("A2:C4").BorderAround Weight:=xlMedium
End Sub

'設定欄寬
Sub Sample22()
    Columns("A:B").AutoFit
    Columns("A:A").ColumnWidth = Columns("A:A").ColumnWidth + 1
End Sub

'設定顏色
Sub Sample37()
    With Range("A1").Font
         .Color = RGB(255, 0, 0)
         .TintAndShade = 0
    End With
    With Range("B1").Interior
         .ThemeColor = xlThemeColorAccent1
         .TintAndShade = 0.8
    End With
End Sub


'儲存格的排序
Sub Sample38()
    With ActiveSheet.Sort.SortFields
         .Clear
         .Add Key:=Range("B2"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
    End With
    With ActiveSheet.Sort
         .SetRange Range("A1:B13")
         .Header = xlYes
         .Orientation = xlTopToBottom
         .Apply
    End With
End Sub


'以列為單位操作儲存格
Sub Sample63()
    Dim i As Long
    For i = 2 To 5
        If Cells(i, 2) > 70 Then
           Range(Cells(i, 1), Cells(i, 5)).Font.Bold = True
        End If
    Next i
End Sub

'移除空白 (1)
Sub Sample68()
    Dim Source As String
    Source = " Microsoft Excel "
    MsgBox Source & "喔" & vbCrLf & Trim(Source) & "喔"
End Sub


'移除空白 (2)
Sub Sample69()
    Dim Source As String
    Source = " Microsoft Excel "
    MsgBox Source & "喔" & vbCrLf & Replace(Source, " ", "") & "喔"
End Sub


'設定格式化條件 (1)
Sub Sample103()
    With Range("A1:A10").FormatConditions
         .Add Type:=xlCellValue, _
              Operator:=xlBetween, _
              Formula1:="40", Formula2:="60"
         .Item(1).Interior.ColorIndex = 3
    End With
End Sub


'刪除格式化條件
Sub Sample105()
    Range("A1:A10").FormatConditions(1).Delete
    Range("A1:A10").FormatConditions.Delete
End Sub


'製作姓名完全不重覆的清單
Sub Sample166()
    Dim MyData As New Collection, i As Long
    On Error Resume Next
    For i = 2 To 101
        MyData.Add Cells(i, 1), Cells(i, 1)
    Next i
    Range("E2") = MyData.Count
    For i = 1 To MyData.Count
         Cells(i + 2, "E") = MyData(i)
    Next i
End Sub


'設定儲存格的背景色.
Sub Sample184()
    Range("B2").Interior.ColorIndex = 3
End Sub


'利用 RGB 函數設定儲存格的背景色
Sub Sample185()
    Range("B2").Interior.Color = RGB(255, 0, 0)
End Sub

'操作存有巨集的活頁簿
Sub Sample2()
    Workbooks.Open Filename:="C:\Book1.xlsx"
     ActiveWorkbook.Sheets(1).Range("A1").Copy _
                    ThisWorkbook.Sheets(1).Range("B1")
End Sub


'隱藏工作表 (1)
Sub Sample20()
    ActiveSheet.Visible = False
End Sub


'重新顯示工作表
Sub Sample21()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Visible = False Then
           ws.Visible = True
        End If
    Next ws
End Sub


'隱藏工作表 (2) - 無法取消隱藏
Sub Sample22()
    ActiveSheet.Visible = xlSheetVeryHidden
End Sub



'保護工作表
Sub Sample23()
    ActiveSheet.Protect Password:="1234"
End Sub

'解除工作表的保護
Sub Sample24()
    ActiveSheet.Unprotect Password:="1234"
End Sub


'在選取儲存格之後自動執行巨集
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        MsgBox "第" & Target.Row & "列" & vbCrLf & _
        "第" & Target.Column & "欄" & vbCrLf & _
        "被選取了"
End Sub


'在儲存格按下滑鼠右鍵時執行巨集
	'在E3:E5範圍之外按下右鍵為一般右鍵
	'在範圍內按下則複製上面一個的東西
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel _
       As Boolean)
    If Application.Intersect(Target, Range("E3:E5")) Is Nothing Then
       Cancel = False
    Else
       Cancel = True
       Target.Offset(-1, 0).Copy Target
    End If
End Sub

'雙按儲存格時執行巨集
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel _
       As Boolean)
    If Application.Intersect(Target, Range("A1:E5")) Is Nothing Then
       Cancel = False
    Else
       Cancel = True
       Target.CurrentRegion.Borders.LineStyle = True
    End If
End Sub



'在「快顯功能表」裡新增命令 - 右鍵選單裡加東西
Sub Sample1()
    With CommandBars("Cell").Controls.Add
         .Caption = "Macro1"
    End With
End Sub

'以「Esc」停止巨集
Sub Sample22()
    Dim i As Long
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo myError
    For i = 1 To 100000
        Application.StatusBar = i
    Next i
    Application.StatusBar = False
    Exit Sub
myError:
    MsgBox "按下 Esc 鍵了"
    Application.StatusBar = False
End Sub


'於「即時運算」視窗裡輸出訊息
Sub Sample29()
    Dim i As Long
    For i = 1 To 10
        Debug.Print i & " - " & Cells(i, 1)
    Next i
End Sub







