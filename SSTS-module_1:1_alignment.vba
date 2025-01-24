Sub FilterModulesFromWorksheet()

    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim inputLastRow As Long
    Dim resultRow As Long
    Dim colHeaders As Range
    Dim moduleName As String
    Dim header As Range
    Dim colIndex As Integer
    Dim cell As Range

    ' 입력 데이터와 결과 저장 워크시트 설정
    Set wsInput = ThisWorkbook.Sheets(1) ' 모듈명을 입력한 워크시트 (Sheet1)
    Set wsData = ThisWorkbook.Sheets(2) ' 원본 데이터가 있는 워크시트 (Sheet2)

    ' 마지막 행 찾기 (데이터 워크시트의 데이터 범위)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' 입력된 모듈명 찾기
    inputLastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    ' 열 헤더 범위 설정
    Set colHeaders = wsData.Rows(1)

    ' 결과를 저장할 새로운 워크시트 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsOutput = ThisWorkbook.Sheets("FilteredResults")
    If Not wsOutput Is Nothing Then wsOutput.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsOutput = ThisWorkbook.Sheets.Add
    wsOutput.Name = "FilteredResults"

    ' 결과 워크시트에 헤더 추가
    wsOutput.Cells(1, 1).Value = "SSTS"
    wsOutput.Cells(1, 2).Value = "Module"
    resultRow = 2

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
                        wsOutput.Cells(resultRow, 1).Value = wsData.Cells(dataCell.Row, 1).Value ' SSTS 값 (첫 번째 열)
                        wsOutput.Cells(resultRow, 2).Value = moduleName
                        resultRow = resultRow + 1
                    End If
                Next dataCell
            Else
                MsgBox "모듈 '" & moduleName & "'은 데이터에 존재하지 않습니다.", vbExclamation
            End If
        End If
    Next cell

    ' 결과 정렬
    If resultRow > 2 Then
        With wsOutput.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsOutput.Columns(1), Order:=xlAscending ' SSTS 정렬
            .SortFields.Add Key:=wsOutput.Columns(2), Order:=xlAscending ' Module 정렬
            .SetRange wsOutput.Range("A1:B" & resultRow - 1)
            .Header = xlYes
            .Apply
        End With
    Else
        MsgBox "일치하는 데이터가 없습니다.", vbInformation
    End If

    MsgBox "필터링 및 정렬이 완료되었습니다.", vbInformation

End Sub
