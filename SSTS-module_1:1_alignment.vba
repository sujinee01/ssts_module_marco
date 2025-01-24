Sub FilterModules()

    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim colHeaders As Range
    Dim colIndex As Integer
    Dim selectedModules As Variant
    Dim moduleName As String
    Dim resultRow As Long
    Dim header As Range
    Dim headerCell As Range
    Dim i As Long

    ' 데이터가 있는 워크시트 지정
    Set ws = ThisWorkbook.Sheets(2)

    ' 열 헤더 범위 설정
    Set colHeaders = ws.Rows(1)

    ' 열 이름의 공백 제거
    For Each headerCell In colHeaders
        headerCell.Value = Trim(headerCell.Value)
    Next headerCell

    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 사용자에게 모듈 입력 요청
    selectedModules = Split(Application.InputBox("모듈을 선택하세요 (예: A,B,C):", "모듈 선택"), ",")

    ' 결과를 저장할 새로운 워크시트 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    Set newWs = ThisWorkbook.Sheets("FilteredResults")
    If Not newWs Is Nothing Then newWs.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "FilteredResults"

    ' 결과 워크시트에 헤더 추가
    newWs.Cells(1, 1).Value = "SSTS"
    newWs.Cells(1, 2).Value = "Module"
    resultRow = 2

    ' 선택된 모듈에 대해 필터링
    For i = LBound(selectedModules) To UBound(selectedModules)
        moduleName = Trim(selectedModules(i))

        ' 모듈이 데이터의 열에 있는지 확인
        Set header = colHeaders.Find(What:=moduleName, LookIn:=xlValues, LookAt:=xlWhole)
        If Not header Is Nothing Then
            colIndex = header.Column

            ' 해당 모듈에서 "x"를 찾고 결과 저장
            For Each cell In ws.Columns(colIndex).Cells(2, 1).Resize(lastRow - 1)
                If cell.Value = "x" Then
                    newWs.Cells(resultRow, 1).Value = ws.Cells(cell.Row, 1).Value ' SSTS 값 (첫 번째 열)
                    newWs.Cells(resultRow, 2).Value = moduleName
                    resultRow = resultRow + 1
                End If
            Next cell
        Else
            MsgBox "모듈 '" & moduleName & "'은 데이터에 존재하지 않습니다.", vbExclamation
        End If
    Next i

    ' 결과 정렬
    If resultRow > 2 Then
        ' 데이터 형식 문자열로 변환
        Dim sortCell As Range
        For Each sortCell In newWs.Range("A2:B" & resultRow - 1)
            sortCell.Value = CStr(sortCell.Value)
        Next sortCell

        ' 정렬 수행
        With newWs.Sort
            .SortFields.Clear
            .SortFields.Add Key:=newWs.Columns(1), Order:=xlAscending ' SSTS 정렬
            .SortFields.Add Key:=newWs.Columns(2), Order:=xlAscending ' Module 정렬
            .SetRange newWs.Range("A1:B" & resultRow - 1)
            .Header = xlYes
            .Apply
        End With
    Else
        MsgBox "일치하는 데이터가 없습니다.", vbInformation
    End If

    MsgBox "필터링 및 정렬이 완료되었습니다.", vbInformation

End Sub
