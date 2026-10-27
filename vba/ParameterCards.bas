Attribute VB_Name = "ParameterCards"
Option Explicit

' Self-contained Excel 2010-2021 module. No Windows API or extra references.
Private Const PREVIEW_PREFIX As String = "ps_preview_"
Private Const PREVIEW_MARKER As String = "PS_PREVIEW|"
Private Const HEIGHT_MARKER As String = "PS_ORIGINAL_HEIGHT_TWIPS="
Private Const BUTTON_NAME As String = "ps_refresh_button"
Private Const PADDING_PT As Double = 3#  ' 4 pixels at 96 dpi.
Private Const MAX_ROW_HEIGHT_PT As Double = 180#
Private Const MIN_ROW_HEIGHT_PT As Double = 24#

Private mUpdating As Boolean
Private mLastSignature As String

Public Sub InstallParameterCards()
    Dim ws As Worksheet
    Dim button As Shape
    Dim anchor As Range
    Dim description As String

    If mUpdating Then Exit Sub
    mUpdating = True
    On Error GoTo InstallFailed
    Set ws = SearchWorksheet()
    Set anchor = ws.Range("K9")
    Set button = ShapeNamed(ws, BUTTON_NAME)

    If button Is Nothing Then
        Set button = ws.Shapes.AddFormControl(xlButtonControl, _
            anchor.Left, anchor.Top, anchor.Width, 22#)
        button.Name = BUTTON_NAME
    Else
        If button.Type <> 8 Then
            Err.Raise vbObjectError + 2100, "ParameterCards", _
                "Имя ps_refresh_button уже занято другим объектом. Переименуйте этот объект."
        End If
        If button.FormControlType <> xlButtonControl Then
            Err.Raise vbObjectError + 2101, "ParameterCards", _
                "Объект ps_refresh_button не является кнопкой формы."
        End If
    End If

    With button
        .Left = anchor.Left
        .Top = anchor.Top
        .Width = anchor.Width
        .Height = 22#
        .Placement = xlMove
        .AlternativeText = "PS_REFRESH_BUTTON"
        .TextFrame.Characters.Text = "Обновить схемы"
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!UpdateParameterImages"
    End With
    mUpdating = False
    UpdateParameterImages
    Exit Sub

InstallFailed:
    description = Err.Description
    mUpdating = False
    On Error Resume Next
    If Not ws Is Nothing Then WriteStatus ws, "Не удалось настроить кнопку: " & description
    On Error GoTo 0
    MsgBox "Не удалось настроить обновление схем." & vbCrLf & description, _
        vbExclamation, "Карточки параметров"
End Sub

Public Sub UpdateParameterImages()
    RefreshImages True, True
End Sub

Public Sub EnsureParameterImages()
    RefreshImages False, False
End Sub

Private Sub RefreshImages(ByVal force As Boolean, ByVal showErrors As Boolean)
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation
    Dim previousBook As Workbook
    Dim previousSheet As Object
    Dim previousSelection As Object
    Dim statesSaved As Boolean
    Dim ws As Worksheet
    Dim tags As Range
    Dim galleries As Range
    Dim indexes As Range
    Dim area As Range
    Dim prerequisites As Range
    Dim sourceShape As Shape
    Dim signature As String
    Dim tag As String
    Dim gallery As String
    Dim description As String
    Dim i As Long
    Dim expected As Long
    Dim created As Long
    Dim missing As Long
    Dim resultCount As Long

    If mUpdating Then Exit Sub
    mUpdating = True
    On Error GoTo RefreshFailed

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    statesSaved = True
    Set previousBook = Application.ActiveWorkbook
    Set previousSheet = Application.ActiveSheet
    Set previousSelection = Application.Selection

    Set ws = SearchWorksheet()
    Set tags = BookRange("Search_ResultCardTag")
    Set galleries = BookRange("Search_ResultCardSheet")
    Set indexes = BookRange("Search_ResultIndex")
    Set area = BookRange("Search_ResultArea")
    ValidateResultRanges ws, area, tags, galleries, indexes

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ' Calculation mode is deliberately not changed. This also works in manual mode.
    Set prerequisites = BookRange("Search_DataBounds")
    prerequisites.Calculate
    Set prerequisites = BookRange("Search_CalcIndex")
    prerequisites.Calculate
    ws.Calculate
    signature = CurrentSignature()
    resultCount = ResultCountValue()

    If Not force Then
        If signature = mLastSignature Then
            If PreviewsComplete(ws, tags, galleries, indexes) Then GoTo RefreshFinished
        End If
    End If

    DeleteOwnPreviews ws, area
    If resultCount = 0 Then
        WriteStatus ws, "По запросу параметры не найдены."
        mLastSignature = signature
        GoTo RefreshFinished
    End If

    ThisWorkbook.Activate
    ws.Activate
    For i = 1 To tags.Rows.Count
        If IsResult(indexes.Cells(i, 1).Value2) Then
            tag = CellText(tags.Cells(i, 1))
            gallery = CellText(galleries.Cells(i, 1))
            If Len(tag) > 0 And Len(gallery) > 0 Then
                expected = expected + 1
                Set sourceShape = FindCardShape(gallery, tag)
                If sourceShape Is Nothing Then
                    missing = missing + 1
                Else
                    AddPreview ws, tags.Cells(i, 1).Row, sourceShape, tag
                    created = created + 1
                End If
            End If
        End If
    Next i

    If CDbl(BookRange("Search_RowCount").Cells(1, 1).Value2) = 0 Then
        WriteStatus ws, "Size не найден в данных. Карточки параметров показаны справочно: " & CStr(created) & "."
    ElseIf expected = 0 Then
        WriteStatus ws, "Для найденных параметров встроенные карточки не назначены."
    ElseIf missing > 0 Then
        WriteStatus ws, "Превью: " & CStr(created) & " из " & CStr(expected) & _
            ". Не найдены исходные карточки: " & CStr(missing) & "."
    Else
        WriteStatus ws, "Схемы соответствуют запросу. Превью: " & CStr(created) & "."
    End If
    mLastSignature = signature
    GoTo RefreshFinished

RefreshFailed:
    description = Err.Description
    mLastSignature = vbNullString
    On Error Resume Next
    If Not ws Is Nothing Then WriteStatus ws, "Ошибка обновления схем: " & description
    On Error GoTo 0
    If showErrors Then
        MsgBox "Не удалось обновить схемы." & vbCrLf & description & vbCrLf & _
            "Поиск и расчёты книги не удалены. Повторите запуск кнопкой после устранения причины.", _
            vbExclamation, "Карточки параметров"
    End If

RefreshFinished:
    ' Restore state even when name lookup, calculation, copying or pasting failed.
    On Error Resume Next
    If statesSaved Then
        Application.CutCopyMode = False
        If Not previousBook Is Nothing Then previousBook.Activate
        If Not previousSheet Is Nothing Then previousSheet.Activate
        If Not previousSelection Is Nothing Then previousSelection.Select
        If Application.Calculation <> previousCalculation Then _
            Application.Calculation = previousCalculation
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEvents
    End If
    mUpdating = False
    On Error GoTo 0
End Sub

Private Function BookRange(ByVal rangeName As String) As Range
    On Error GoTo MissingName
    Set BookRange = ThisWorkbook.Names(rangeName).RefersToRange
    If BookRange Is Nothing Then GoTo MissingName
    Exit Function
MissingName:
    Err.Raise vbObjectError + 2110, "ParameterCards", _
        "В книге отсутствует корректный именованный диапазон " & rangeName & "."
End Function

Private Function SearchWorksheet() As Worksheet
    Set SearchWorksheet = BookRange("Search_Source").Worksheet
End Function

Private Sub ValidateResultRanges(ByVal ws As Worksheet, ByVal area As Range, _
    ByVal tags As Range, ByVal galleries As Range, ByVal indexes As Range)
    If Not area.Worksheet Is ws Then GoTo InvalidRanges
    If Not tags.Worksheet Is ws Then GoTo InvalidRanges
    If Not galleries.Worksheet Is ws Then GoTo InvalidRanges
    If Not indexes.Worksheet Is ws Then GoTo InvalidRanges
    If tags.Columns.Count <> 1 Or galleries.Columns.Count <> 1 Or _
        indexes.Columns.Count <> 1 Then GoTo InvalidRanges
    If tags.Row <> area.Row Or galleries.Row <> area.Row Or _
        indexes.Row <> area.Row Then GoTo InvalidRanges
    If tags.Rows.Count <> area.Rows.Count Or galleries.Rows.Count <> area.Rows.Count Or _
        indexes.Rows.Count <> area.Rows.Count Then GoTo InvalidRanges
    Exit Sub
InvalidRanges:
    Err.Raise vbObjectError + 2111, "ParameterCards", _
        "Именованные диапазоны результатов относятся к разным листам или строкам."
End Sub

Private Function CellText(ByVal cell As Range) As String
    Dim value As Variant
    value = cell.Cells(1, 1).Value2
    If IsError(value) Then
        Err.Raise vbObjectError + 2112, "ParameterCards", _
            "Ошибка формулы в ячейке " & cell.Address(False, False) & "."
    End If
    CellText = CStr(value)
End Function

Private Function SignaturePart(ByVal rangeName As String) As String
    Dim value As String
    value = CellText(BookRange(rangeName))
    SignaturePart = CStr(Len(value)) & ":" & value & ";"
End Function

Private Function CurrentSignature() As String
    CurrentSignature = SignaturePart("Search_Source") & _
        SignaturePart("Search_Size") & SignaturePart("Search_Filter") & _
        SignaturePart("Search_Count") & SignaturePart("Search_RowCount")
End Function

Private Function ResultCountValue() As Long
    Dim value As Variant
    value = BookRange("Search_Count").Cells(1, 1).Value2
    If IsError(value) Then
        Err.Raise vbObjectError + 2113, "ParameterCards", _
            "Не удалось прочитать количество найденных параметров."
    End If
    If Not IsNumeric(value) Then
        Err.Raise vbObjectError + 2113, "ParameterCards", _
            "Не удалось прочитать количество найденных параметров."
    End If
    If CDbl(value) < 0 Then
        Err.Raise vbObjectError + 2114, "ParameterCards", _
            "Количество найденных параметров не может быть отрицательным."
    End If
    ResultCountValue = CLng(value)
End Function

Private Function IsResult(ByVal value As Variant) As Boolean
    If IsError(value) Then
        Err.Raise vbObjectError + 2115, "ParameterCards", _
            "В индексе результатов обнаружена ошибка формулы."
    End If
    If IsNumeric(value) Then IsResult = (CDbl(value) > 0)
End Function

Private Function ShapeNamed(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set ShapeNamed = ws.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Function FindCardShape(ByVal gallery As String, ByVal tag As String) As Shape
    Dim source As Worksheet
    Dim item As Shape
    On Error Resume Next
    Set source = ThisWorkbook.Worksheets(gallery)
    On Error GoTo 0
    If source Is Nothing Then Exit Function
    For Each item In source.Shapes
        If StrComp(item.AlternativeText, tag, vbBinaryCompare) = 0 Then
            Set FindCardShape = item
            Exit Function
        End If
    Next item
End Function

Private Function IsOwnPreview(ByVal shapeName As String) As Boolean
    IsOwnPreview = (StrComp(Left$(shapeName, Len(PREVIEW_PREFIX)), _
        PREVIEW_PREFIX, vbBinaryCompare) = 0)
End Function

Private Function PreviewsComplete(ByVal ws As Worksheet, ByVal tags As Range, _
    ByVal galleries As Range, ByVal indexes As Range) As Boolean
    Dim expected As Long
    Dim actual As Long
    Dim i As Long
    Dim tag As String
    Dim gallery As String
    Dim item As Shape
    Dim marker As String
    For i = 1 To tags.Rows.Count
        If IsResult(indexes.Cells(i, 1).Value2) Then
            tag = CellText(tags.Cells(i, 1))
            gallery = CellText(galleries.Cells(i, 1))
            If Len(tag) > 0 And Len(gallery) > 0 Then
                expected = expected + 1
                Set item = ShapeNamed(ws, PREVIEW_PREFIX & CStr(tags.Cells(i, 1).Row))
                If item Is Nothing Then Exit Function
                marker = PREVIEW_MARKER & tag & vbLf
                If StrComp(Left$(item.AlternativeText, Len(marker)), marker, _
                    vbBinaryCompare) <> 0 Then Exit Function
            End If
        End If
    Next i
    For Each item In ws.Shapes
        If IsOwnPreview(item.Name) Then actual = actual + 1
    Next item
    PreviewsComplete = (actual = expected)
End Function

Private Sub RestorePreviewRowHeight(ByVal ws As Worksheet, ByVal item As Shape, ByVal area As Range)
    Dim suffix As String
    Dim markerPosition As Long
    Dim heightTwips As String
    Dim rowNumber As Long
    Dim height As Double
    suffix = Mid$(item.Name, Len(PREVIEW_PREFIX) + 1)
    If Not IsNumeric(suffix) Then Exit Sub
    rowNumber = CLng(suffix)
    If rowNumber < area.Row Or rowNumber >= area.Row + area.Rows.Count Then Exit Sub
    markerPosition = InStr(1, item.AlternativeText, vbLf & HEIGHT_MARKER, vbBinaryCompare)
    If markerPosition = 0 Then Exit Sub
    heightTwips = Mid$(item.AlternativeText, markerPosition + 1 + Len(HEIGHT_MARKER))
    If Not IsNumeric(heightTwips) Then Exit Sub
    height = CDbl(heightTwips) / 20#
    If height > 0 And height <= 409.5 Then ws.Rows(rowNumber).RowHeight = height
End Sub

Private Sub DeleteOwnPreviews(ByVal ws As Worksheet, ByVal area As Range)
    Dim i As Long
    Dim item As Shape
    For i = ws.Shapes.Count To 1 Step -1
        Set item = ws.Shapes(i)
        If IsOwnPreview(item.Name) Then
            RestorePreviewRowHeight ws, item, area
            item.Delete
        End If
    Next i
End Sub

Private Sub AddPreview(ByVal ws As Worksheet, ByVal rowNumber As Long, _
    ByVal source As Shape, ByVal tag As String)
    Dim picture As Shape
    Dim target As Range
    Dim oldHeight As Double
    Dim originalWidth As Double
    Dim originalHeight As Double
    Dim scaleFactor As Double
    Dim availableWidth As Double
    Dim availableHeight As Double
    Dim newHeight As Double
    Dim beforeCount As Long
    Dim number As Long
    Dim description As String

    On Error GoTo PasteFailed
    Set target = ws.Cells(rowNumber, 7)
    oldHeight = ws.Rows(rowNumber).RowHeight
    availableWidth = ws.Columns(7).Width - 2# * PADDING_PT
    availableHeight = MAX_ROW_HEIGHT_PT - 2# * PADDING_PT
    If availableWidth <= 0 Then
        Err.Raise vbObjectError + 2120, "ParameterCards", "Столбец G слишком узкий для превью."
    End If
    beforeCount = ws.Shapes.Count
    source.Copy
    ws.Paste
    If ws.Shapes.Count <> beforeCount + 1 Then
        Err.Raise vbObjectError + 2121, "ParameterCards", "Excel не вставил одну картинку из буфера обмена."
    End If
    Set picture = ws.Shapes(ws.Shapes.Count)
    picture.Name = PREVIEW_PREFIX & CStr(rowNumber)
    picture.AlternativeText = PREVIEW_MARKER & tag & vbLf & _
        HEIGHT_MARKER & CStr(CLng(oldHeight * 20#))
    originalWidth = picture.Width
    originalHeight = picture.Height
    If originalWidth <= 0 Or originalHeight <= 0 Then
        Err.Raise vbObjectError + 2122, "ParameterCards", "Некорректный размер исходной карточки."
    End If
    scaleFactor = availableWidth / originalWidth
    If originalHeight * scaleFactor > availableHeight Then scaleFactor = availableHeight / originalHeight
    newHeight = originalHeight * scaleFactor + 2# * PADDING_PT
    If newHeight < oldHeight Then newHeight = oldHeight
    If newHeight < MIN_ROW_HEIGHT_PT Then newHeight = MIN_ROW_HEIGHT_PT
    If newHeight > MAX_ROW_HEIGHT_PT Then newHeight = MAX_ROW_HEIGHT_PT
    ws.Rows(rowNumber).RowHeight = newHeight
    With picture
        .LockAspectRatio = 0
        .Width = originalWidth * scaleFactor
        .Height = originalHeight * scaleFactor
        .LockAspectRatio = -1
        .Placement = xlMoveAndSize
        .Left = target.Left + PADDING_PT
        .Top = target.Top + (ws.Rows(rowNumber).RowHeight - .Height) / 2#
    End With
    Exit Sub

PasteFailed:
    number = Err.Number
    description = Err.Description
    On Error Resume Next
    If Not picture Is Nothing Then
        If IsOwnPreview(picture.Name) Then picture.Delete
    End If
    If oldHeight > 0 Then ws.Rows(rowNumber).RowHeight = oldHeight
    On Error GoTo 0
    Err.Raise number, "ParameterCards.AddPreview", description
End Sub

Private Sub WriteStatus(ByVal ws As Worksheet, ByVal message As String)
    Dim target As Range
    Set target = BookRange("Search_ImageStatus")
    If Not target.Worksheet Is ws Then
        Err.Raise vbObjectError + 2123, "ParameterCards", _
            "Диапазон статуса находится на другом листе."
    End If
    target.Cells(1, 1).Value2 = message
End Sub
