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

    ' 입력 데이터를 포함한 워크시트 지정 (Sheet1)
    Set wsInput = ThisWorkbook.Sheets(1)
    ' 원본 데이터를 포함한 워크시트 지정 (Sheet2)
    Set wsData = ThisWorkbook.Sheets(2)

    ' 마지막 행 찾기 (데이터 워크시트의 데이터 범위)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' 입력된 모듈명의 마지막 행 찾기
    inputLastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    ' 열 헤더 범위 설정
    Set colHeaders = wsData.Rows(1)

    ' 결과 표시를 위한 시작 행 설정 (입력 데이터 옆으로 표시)
    resultRow = 2 ' 결과는 B열부터 시작한다고 가정

    ' 기존 결과 제거 (B2:C 영역 초기화)
    wsInput.Range("B2:C" & wsInput.Rows.Count).ClearContents

    ' 입력된 모듈명을 기준으로 필터링
    For Each cell In wsInput.Range("A2:A" & inputLastRow)
        moduleName = Trim(cell.Value)
        If moduleName <> "" Then
            ' 모듈이 데이터의 열에 있는지 확인
            Set header = colHeaders.Find(What:=moduleName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not header Is Nothing Then
                colIndex = header.Column

                ' 해당 모듈에서 "x"를 찾고 결과 저장
                For Each dataCell In wsData.Columns(colIndex).Cells(2, 1).Resize(lastRow - 1)
                    If dataCell.Value = "x" Then
                        wsInput.Cells(resultRow, 2).Value = wsData.Cells(dataCell.Row, 1).Value ' SSTS 값
                        wsInput.Cells(resultRow, 3).Value = moduleName ' 모듈명
                        resultRow = resultRow + 1
                    End If
                Next dataCell
            Else
                MsgBox "모듈 '" & moduleName & "'은 데이터에 존재하지 않습니다.", vbExclamation
            End If
        End If
    Next cell

    ' 결과가 없는 경우 메시지 표시
    If resultRow = 2 Then
        MsgBox "일치하는 데이터가 없습니다.", vbInformation
    Else
        MsgBox "필터링 및 결과 표시가 완료되었습니다.", vbInformation
    End If

End Sub
