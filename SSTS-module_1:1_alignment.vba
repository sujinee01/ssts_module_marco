Sub FilterModulesToSameWorksheet()

    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim inputLastRow As Long
    Dim resultRow As Long
    Dim colHeaders As Range
    Dim moduleName As String
    Dim header As Range
    Dim colIndex As Integer
    Dim cell As Range

    ' 입력 데이터를 포함한 워크시트 지정 (Sheet2)
    Set wsInput = ThisWorkbook.Sheets(2)
    ' 원본 데이터를 포함한 워크시트 지정 (Sheet3)
    Set wsData = ThisWorkbook.Sheets(3)

    ' 원본 데이터의 마지막 행 찾기
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' 입력된 모듈명의 마지막 행 찾기
    inputLastRow = wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).Row

    ' 원본 데이터의 열 헤더 범위 설정
    Set colHeaders = wsData.Rows(1)

    ' 결과 표시를 위한 시작 행 설정
    resultRow = 3 ' 결과 데이터는 3행부터 시작하도록 설정

    ' 기존 결과 제거 (D1:E 영역 초기화)
    wsInput.Range("C1:Z" & wsInput.Rows.Count).ClearContents

    ' 결과 워크시트에 헤더 추가
    wsInput.Cells(2, 2).Value = "찾고 싶은 Module"
    wsInput.Cells(2, 4).Value = "SSTS"
    wsInput.Cells(2, 5).Value = "Module"

    ' 입력된 모듈명을 기준으로 필터링
    For Each cell In wsInput.Range("B3:B" & inputLastRow) ' 2번 시트 B열에 입력된 모듈명 반복
        moduleName = Trim(cell.Value)
        If moduleName <> "" Then
            ' 원본 데이터에서 모듈이 존재하는 열 확인
            Set header = colHeaders.Find(What:=moduleName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not header Is Nothing Then
                colIndex = header.Column

                ' 해당 모듈에서 "x"를 찾고 결과 저장
                For Each dataCell In wsData.Columns(colIndex).Cells(2, 1).Resize(lastRow - 1)
                    If dataCell.Value = "x" Then
                        wsInput.Cells(resultRow, 4).Value = wsData.Cells(dataCell.Row, 1).Value ' SSTS 값 (A열)
                        wsInput.Cells(resultRow, 5).Value = moduleName ' 모듈명
                        resultRow = resultRow + 1
                    End If
                Next dataCell
            Else
                MsgBox "모듈 '" & moduleName & "'은 원본 데이터에 존재하지 않습니다.", vbExclamation
            End If
        End If
    Next cell

    ' 결과 정렬 (SSTS 기준으로 오름차순)
    If resultRow > 3 Then
        With wsInput.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsInput.Columns(4), Order:=xlAscending ' SSTS 정렬 (D열)
            .SetRange wsInput.Range("D2:E" & resultRow - 1) ' 정렬 범위
            .header = xlYes
            .Apply
        End With
        MsgBox "필터링 및 정렬이 완료되었습니다.", vbInformation
    Else
        MsgBox "일치하는 데이터가 없습니다.", vbInformation
    End If

End Sub

