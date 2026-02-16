Option Explicit

'==============================
' ПЕРЕМЕННЫЕ
'==============================
Private srcRange As Range
Private srcSheet As Worksheet
Private srcWb As Workbook
Private destCell As Range

'==============================
' ИНИЦИАЛИЗАЦИЯ ФОРМЫ
'==============================
Private Sub UserForm_Initialize()
    ' Проверяем, что есть выделение
    If TypeName(Selection) <> "Range" Then
        MsgBox "Перед запуском выделите строки для копирования!", vbExclamation
        Unload Me
        Exit Sub
    End If

    ' Сохраняем источник
    Set srcRange = Selection
    Set srcSheet = srcRange.Worksheet
    Set srcWb = srcSheet.Parent

    ' Настройка ListBox
    lstColumns.ColumnCount = 2
    lstColumns.ColumnWidths = "70;70"
    lstColumns.Height = 140

    ' Загружаем книги и сохранённые настройки
    LoadWorkbooks
    LoadSavedSettings
End Sub

'==============================
' ЗАГРУЗКА ОТКРЫТЫХ КНИГ
'==============================
Private Sub LoadWorkbooks()
    Dim wb As Workbook

    cmbTargetWorkbook.Clear
    cmbTargetSheet.Clear

    For Each wb In Application.Workbooks
        If wb.Name <> srcWb.Name Then
            cmbTargetWorkbook.AddItem wb.Name
        End If
    Next wb

    If cmbTargetWorkbook.ListCount > 0 Then
        cmbTargetWorkbook.ListIndex = 0
    End If
End Sub

Private Sub cmbTargetWorkbook_Change()
    Dim ws As Worksheet

    cmbTargetSheet.Clear

    If cmbTargetWorkbook.Value = "" Then Exit Sub
    If Not WorkbookExists(cmbTargetWorkbook.Value) Then Exit Sub

    For Each ws In Workbooks(cmbTargetWorkbook.Value).Worksheets
        cmbTargetSheet.AddItem ws.Name
    Next ws

    If cmbTargetSheet.ListCount > 0 Then
        cmbTargetSheet.ListIndex = 0
    End If
End Sub

'==============================
' ВАЛИДАЦИЯ
'==============================
Private Function WorkbookExists(ByVal wbName As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, wbName, vbTextCompare) = 0 Then
            WorkbookExists = True
            Exit Function
        End If
    Next wb
End Function

Private Function IsValidColumnLetter(ByVal colLetter As String) As Boolean
    Dim i As Long
    Dim ch As Integer

    colLetter = UCase$(Trim$(colLetter))
    If colLetter = "" Then Exit Function

    For i = 1 To Len(colLetter)
        ch = Asc(Mid$(colLetter, i, 1))
        If ch < 65 Or ch > 90 Then Exit Function
    Next i

    On Error Resume Next
    IsValidColumnLetter = (ColLetterToNumber(colLetter) > 0)
    On Error GoTo 0
End Function

'==============================
' ДОБАВЛЕНИЕ ПАРЫ СТОЛБЦОВ (буквы)
'==============================
Private Function ColLetterToNumber(ByVal colLetter As String) As Long
    ColLetterToNumber = srcSheet.Range(UCase$(Trim$(colLetter)) & "1").Column
End Function

Private Sub btnAddPair_Click()
    Dim srcLetter As String
    Dim destLetter As String

    srcLetter = UCase$(Trim$(txtSrcCol.Value))
    destLetter = UCase$(Trim$(txtDestCol.Value))

    If srcLetter = "" Or destLetter = "" Then Exit Sub

    If Not IsValidColumnLetter(srcLetter) Or Not IsValidColumnLetter(destLetter) Then
        MsgBox "Введите корректные буквы столбцов (например A, B, AA).", vbExclamation
        Exit Sub
    End If

    lstColumns.AddItem srcLetter
    lstColumns.List(lstColumns.ListCount - 1, 1) = destLetter

    txtSrcCol = ""
    txtDestCol = ""
End Sub

Private Sub btnRemovePair_Click()
    If lstColumns.ListIndex >= 0 Then
        lstColumns.RemoveItem lstColumns.ListIndex
    End If
End Sub

'==============================
' ВЫБОР СТАРТОВОЙ ЯЧЕЙКИ
'==============================
Private Sub btnSelectCell_Click()
    Dim tmp As Range

    Me.Hide
    On Error Resume Next
    Set tmp = Application.InputBox( _
        Prompt:="Выберите стартовую ячейку", _
        Title:="Выбор ячейки", _
        Type:=8)
    On Error GoTo 0

    If Not tmp Is Nothing Then Set destCell = tmp
    Me.Show
End Sub

'==============================
' ЗАПУСК ПЕРЕНОСА
'==============================
Private Sub btnRun_Click()
    Dim destWb As Workbook
    Dim destSheet As Worksheet
    Dim i As Long
    Dim j As Long
    Dim destRow As Long
    Dim srcCol As Long
    Dim destCol As Long

    If cmbTargetWorkbook.Value = "" Or cmbTargetSheet.Value = "" Then
        MsgBox "Выберите книгу и лист назначения!", vbExclamation
        Exit Sub
    End If

    If Not WorkbookExists(cmbTargetWorkbook.Value) Then
        MsgBox "Книга назначения не найдена среди открытых.", vbExclamation
        Exit Sub
    End If

    If destCell Is Nothing Then
        MsgBox "Не выбрана ячейка назначения!", vbExclamation
        Exit Sub
    End If

    If lstColumns.ListCount = 0 Then
        MsgBox "Нет пар столбцов!", vbExclamation
        Exit Sub
    End If

    Set destWb = Workbooks(cmbTargetWorkbook.Value)
    Set destSheet = destWb.Worksheets(cmbTargetSheet.Value)

    If Not destCell.Worksheet Is destSheet Then
        MsgBox "Стартовая ячейка должна быть на выбранном листе назначения.", vbExclamation
        Exit Sub
    End If

    destRow = destCell.Row
    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    For i = 1 To srcRange.Rows.Count
        For j = 0 To lstColumns.ListCount - 1
            srcCol = ColLetterToNumber(lstColumns.List(j, 0))
            destCol = ColLetterToNumber(lstColumns.List(j, 1))

            destSheet.Cells(destRow, destCol).Value = _
                srcSheet.Cells(srcRange.Row + i - 1, srcCol).Value
        Next j
        destRow = destRow + 1
    Next i

    Application.ScreenUpdating = True

    MsgBox "Перенос выполнен!", vbInformation
    Unload Me
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при переносе: " & Err.Description, vbExclamation
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

'==============================
' СОХРАНЕНИЕ И ЗАГРУЗКА НАСТРОЕК
'==============================
Private Sub btnSaveSetting_Click()
    Dim i As Long
    Dim settingStr As String
    Dim nameSetting As String

    nameSetting = Trim$(InputBox("Введите имя для этой настройки:", "Сохранить настройку"))
    If nameSetting = "" Then Exit Sub

    If cmbTargetWorkbook.Value = "" Or cmbTargetSheet.Value = "" Then
        MsgBox "Перед сохранением выберите книгу и лист назначения.", vbExclamation
        Exit Sub
    End If

    ' Собираем текущие настройки
    settingStr = cmbTargetWorkbook.Value & "|" & cmbTargetSheet.Value & "|"
    For i = 0 To lstColumns.ListCount - 1
        settingStr = settingStr & lstColumns.List(i, 0) & ":" & lstColumns.List(i, 1) & ";"
    Next i

    ' Сохраняем настройку
    SaveSetting "CopyTool", "SavedSettings", nameSetting, settingStr
    LoadSavedSettings
    cmbSaved.Value = nameSetting
    MsgBox "Настройка сохранена!", vbInformation
End Sub

Private Sub LoadSavedSettings()
    Dim settings As Variant
    Dim i As Long

    cmbSaved.Clear
    On Error Resume Next
    settings = GetAllSettings("CopyTool", "SavedSettings")
    On Error GoTo 0

    If IsArray(settings) Then
        For i = LBound(settings, 1) To UBound(settings, 1)
            cmbSaved.AddItem CStr(settings(i, 0))
        Next i
    End If
End Sub

Private Sub btnLoadSetting_Click()
    Dim parts() As String
    Dim pairs() As String
    Dim pairParts() As String
    Dim i As Long
    Dim selectedSetting As String
    Dim rawSetting As String

    selectedSetting = cmbSaved.Value
    If selectedSetting = "" Then Exit Sub

    rawSetting = GetSetting("CopyTool", "SavedSettings", selectedSetting, "")
    If rawSetting = "" Then Exit Sub

    parts = Split(rawSetting, "|")
    If UBound(parts) < 2 Then
        MsgBox "Повреждённый формат сохранённой настройки.", vbExclamation
        Exit Sub
    End If

    If WorkbookExists(parts(0)) Then
        cmbTargetWorkbook.Value = parts(0)
        cmbTargetWorkbook_Change
    Else
        MsgBox "Книга из сохранённой настройки не открыта: " & parts(0), vbExclamation
    End If

    If cmbTargetSheet.ListCount > 0 Then
        On Error Resume Next
        cmbTargetSheet.Value = parts(1)
        On Error GoTo 0
    End If

    lstColumns.Clear
    pairs = Split(parts(2), ";")
    For i = 0 To UBound(pairs)
        If pairs(i) <> "" Then
            pairParts = Split(pairs(i), ":")
            If UBound(pairParts) = 1 Then
                If IsValidColumnLetter(pairParts(0)) And IsValidColumnLetter(pairParts(1)) Then
                    lstColumns.AddItem UCase$(Trim$(pairParts(0)))
                    lstColumns.List(lstColumns.ListCount - 1, 1) = UCase$(Trim$(pairParts(1)))
                End If
            End If
        End If
    Next i
End Sub
