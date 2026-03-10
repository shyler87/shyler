Option Explicit

'==============================
' ПЕРЕМЕННЫЕ
'==============================
Private srcRange As Range
Private srcSheet As Worksheet
Private srcWb As Workbook
Private destCell As Range

Private dragIndex As Long
Private editIndex As Long


'==============================
' ИНИЦИАЛИЗАЦИЯ ФОРМЫ
'==============================
Private Sub UserForm_Initialize()

    If TypeName(Selection) <> "Range" Then
        MsgBox "Перед запуском выделите строки для копирования!", vbExclamation
        Unload Me
        Exit Sub
    End If

    Set srcRange = Selection
    Set srcSheet = srcRange.Worksheet
    Set srcWb = srcSheet.Parent

    lstColumns.ColumnCount = 2
    lstColumns.ColumnWidths = "70;70"

    dragIndex = -1
    editIndex = -1

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
' ПРОВЕРКА КНИГИ
'==============================
Private Function WorkbookExists(ByVal wbName As String) As Boolean

    Dim wb As Workbook
    WorkbookExists = False

    For Each wb In Application.Workbooks

        If StrComp(wb.Name, wbName, vbTextCompare) = 0 Then
            WorkbookExists = True
            Exit Function
        End If

    Next wb

End Function


'==============================
' ПРОВЕРКА КОЛОНКИ
'==============================
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
    IsValidColumnLetter = (ColLetterToNumber(srcSheet, colLetter) > 0)
    On Error GoTo 0

End Function


'==============================
' БУКВА -> НОМЕР КОЛОНКИ
'==============================
Private Function ColLetterToNumber(ws As Worksheet, ByVal colLetter As String) As Long

    ColLetterToNumber = ws.Range(UCase$(Trim$(colLetter)) & "1").Column

End Function


'==============================
' ДОБАВИТЬ / ИЗМЕНИТЬ ПАРУ
'==============================
Private Sub btnAddPair_Click()

    Dim srcLetter As String
    Dim destLetter As String

    srcLetter = UCase$(Trim$(txtSrcCol.Value))
    destLetter = UCase$(Trim$(txtDestCol.Value))

    If srcLetter = "" Or destLetter = "" Then Exit Sub

    If Not IsValidColumnLetter(srcLetter) _
    Or Not IsValidColumnLetter(destLetter) Then

        MsgBox "Введите корректные буквы столбцов (A, B, AA...).", vbExclamation
        Exit Sub

    End If

    If editIndex >= 0 And editIndex < lstColumns.ListCount Then

        lstColumns.List(editIndex, 0) = srcLetter
        lstColumns.List(editIndex, 1) = destLetter
        editIndex = -1

    Else

        lstColumns.AddItem srcLetter
        lstColumns.List(lstColumns.ListCount - 1, 1) = destLetter

    End If

    txtSrcCol = ""
    txtDestCol = ""

    txtSrcCol.SetFocus

End Sub


Private Sub btnRemovePair_Click()

    If lstColumns.ListIndex >= 0 Then
        lstColumns.RemoveItem lstColumns.ListIndex
    End If

End Sub


'==============================
' DRAG & DROP
'==============================
Private Sub lstColumns_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    If lstColumns.ListIndex >= 0 Then
        dragIndex = lstColumns.ListIndex
    End If

End Sub


Private Sub lstColumns_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    Dim newIndex As Long
    Dim t1 As String
    Dim t2 As String

    If Button = 1 Then

        newIndex = lstColumns.ListIndex

        If newIndex <> dragIndex And newIndex >= 0 Then

            t1 = lstColumns.List(dragIndex, 0)
            t2 = lstColumns.List(dragIndex, 1)

            lstColumns.List(dragIndex, 0) = lstColumns.List(newIndex, 0)
            lstColumns.List(dragIndex, 1) = lstColumns.List(newIndex, 1)

            lstColumns.List(newIndex, 0) = t1
            lstColumns.List(newIndex, 1) = t2

            dragIndex = newIndex

        End If

    End If

End Sub


'==============================
' ДВОЙНОЙ КЛИК (РЕДАКТИРОВАНИЕ)
'==============================
Private Sub lstColumns_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If lstColumns.ListIndex < 0 Then Exit Sub

    editIndex = lstColumns.ListIndex

    txtSrcCol.Value = lstColumns.List(editIndex, 0)
    txtDestCol.Value = lstColumns.List(editIndex, 1)

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
        Type:=8)

    On Error GoTo 0

    If Not tmp Is Nothing Then
        Set destCell = tmp
    End If

    Me.Show

End Sub


'==============================
' ПЕРЕНОС С УЧЕТОМ ФИЛЬТРА
'==============================
Private Sub btnRun_Click()

    Dim destWb As Workbook
    Dim destSheet As Worksheet

    Dim visibleRows As Range
    Dim r As Range

    Dim j As Long
    Dim destRow As Long

    Dim srcCol As Long
    Dim destCol As Long

    If cmbTargetWorkbook.Value = "" Or cmbTargetSheet.Value = "" Then
        MsgBox "Выберите книгу и лист назначения!", vbExclamation
        Exit Sub
    End If

    If destCell Is Nothing Then
        MsgBox "Выберите стартовую ячейку назначения!", vbExclamation
        Exit Sub
    End If

    If lstColumns.ListCount = 0 Then
        MsgBox "Нет пар столбцов!", vbExclamation
        Exit Sub
    End If

    Set destWb = Workbooks(cmbTargetWorkbook.Value)
    Set destSheet = destWb.Worksheets(cmbTargetSheet.Value)

    On Error Resume Next
    Set visibleRows = srcRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleRows Is Nothing Then
        MsgBox "Нет строк для переноса.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    destRow = destCell.Row

    For Each r In visibleRows.Rows

        For j = 0 To lstColumns.ListCount - 1

            srcCol = ColLetterToNumber(srcSheet, lstColumns.List(j, 0))
            destCol = ColLetterToNumber(destSheet, lstColumns.List(j, 1))

            destSheet.Cells(destRow, destCol).Value = _
                srcSheet.Cells(r.Row, srcCol).Value

        Next j

        destRow = destRow + 1

    Next r

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Перенос выполнен!", vbInformation

End Sub


Private Sub btnCancel_Click()
    Unload Me
End Sub


'==============================
' СОХРАНЕНИЕ НАСТРОЕК
'==============================
Private Sub btnSaveSetting_Click()

    Dim i As Long
    Dim s As String
    Dim nameSetting As String

    nameSetting = Trim$(InputBox("Введите имя настройки"))

    If nameSetting = "" Then Exit Sub

    If cmbTargetWorkbook.Value = "" Or cmbTargetSheet.Value = "" Then
        MsgBox "Выберите книгу и лист.", vbExclamation
        Exit Sub
    End If

    s = cmbTargetWorkbook.Value & "|" & cmbTargetSheet.Value & "|"

    For i = 0 To lstColumns.ListCount - 1

        s = s & _
            lstColumns.List(i, 0) & ":" & _
            lstColumns.List(i, 1) & ";"

    Next i

    SaveSetting "CopyTool", "SavedSettings", nameSetting, s

    LoadSavedSettings

    cmbSaved.Value = nameSetting

    MsgBox "Настройка сохранена!", vbInformation

End Sub


'==============================
' ЗАГРУЗКА СПИСКА НАСТРОЕК
'==============================
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


'==============================
' ЗАГРУЗКА НАСТРОЙКИ
'==============================
Private Sub btnLoadSetting_Click()

    Dim parts As Variant
    Dim pairs As Variant
    Dim p As Variant

    Dim i As Long
    Dim raw As String

    If cmbSaved.Value = "" Then Exit Sub

    raw = GetSetting("CopyTool", "SavedSettings", cmbSaved.Value, "")

    If raw = "" Then Exit Sub

    parts = Split(raw, "|")

    If UBound(parts) < 2 Then
        MsgBox "Ошибка формата сохраненной настройки.", vbExclamation
        Exit Sub
    End If

    cmbTargetWorkbook = parts(0)
    cmbTargetWorkbook_Change
    cmbTargetSheet = parts(1)

    lstColumns.Clear

    pairs = Split(parts(2), ";")

    For i = 0 To UBound(pairs)

        If pairs(i) <> "" Then

            p = Split(pairs(i), ":")

            If UBound(p) = 1 Then
                lstColumns.AddItem p(0)
                lstColumns.List(lstColumns.ListCount - 1, 1) = p(1)
            End If

        End If

    Next i

End Sub
