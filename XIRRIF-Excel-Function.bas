Function XIRRIF( _
        CashflowAmountsColumn As Range, _
        CashflowDatesColumn As Range, _
        Optional FundNamesColumn As Range, _
        Optional FundNameConditionCell As Range, _
        Optional GuessValue As Double)
'
' This function returns XIRR based on a condition
'
    On Error Resume Next
    
    Dim FirstRow As Long, LastRow As Long
    Dim WkSht As Worksheet
    Dim FundStartCell As Range, FundEndCell As Range
    Dim AmountsColumnIndex As Integer, DatesColumnIndex As Integer, FundNamesColumnIndex As Integer
    
    Set WkSht = ActiveSheet
    AmountsColumnIndex = CashflowAmountsColumn.Column
    DatesColumnIndex = CashflowDatesColumn.Column
    FundNamesColumnIndex = FundNamesColumn.Column
    
    With WkSht
    
        FirstRow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
        LastRow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        Set FundStartCell = FundNamesColumn.Find(FundNameConditionCell.Value, SearchOrder:=xlByRows, SearchDirection:=xlNext, LookAt:=xlWhole)
        Set FundEndCell = FundNamesColumn.Find(FundNameConditionCell.Value, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole)
        
        Do While Not IsDate(.Cells(FirstRow, DatesColumnIndex))
            FirstRow = FirstRow + 1
        Loop
        
        If FundNamesColumn Is Nothing And FundNameConditionCell Is Nothing Then
            XIRRIF = _
                WorksheetFunction.Xirr( _
                    .Range(.Cells(FirstRow, AmountsColumnIndex), .Cells(LastRow, AmountsColumnIndex)), _
                    .Range(.Cells(FirstRow, DatesColumnIndex), .Cells(LastRow, DatesColumnIndex)), _
                    GuessValue)
        ElseIf Not FundStartCell Is Nothing And Not FundEndCell Is Nothing Then
            XIRRIF = _
                WorksheetFunction.Xirr( _
                    .Range(.Cells(FundStartCell.Row, AmountsColumnIndex), .Cells(FundEndCell.Row, AmountsColumnIndex)), _
                    .Range(.Cells(FundStartCell.Row, DatesColumnIndex), .Cells(FundEndCell.Row, DatesColumnIndex)), _
                    GuessValue)
        Else
            XIRRIF = CVErr(xlErrNA)
        End If
        
    End With
    
End Function
