Sub Top5Highlight()
    
    'If columns extend past P update the following lines
    Dim lastCol As String 'name of last column
    lastCol = "P"
    Dim latColInt As Integer 'number of last column (A=1, B=2, etc.)
    lastColInt = 16
    
    'Application.ScreenUpdating = False
    Dim dictrng As Range
    Dim cl As Range
    Dim dict As Object
    Dim ky As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    With ActiveSheet
        Set dictrng = .Range(.Range("B2"), .Range("B" & .Rows.Count).End(xlUp))
    End With
    
    For Each cl In dictrng
        If Not dict.exists(cl.Value) Then
            dict.Add cl.Value, cl.Value
        End If
    Next cl
    
    For Each ky In dict.keys
    
        ActiveSheet.Range("A1:" & lastCol & ActiveSheet.Rows.Count).AutoFilter Field:=2, Criteria1:=ky
        
        Dim rng As Range
        Dim cell As Range
        Dim cnt As Long
        Set rng = Range("A2", Range("A65536").End(xlUp)).SpecialCells(xlCellTypeVisible)
        Dim rCell As Double
        rCell = ActiveSheet.AutoFilter.Range.Offset(1, 0).Row 'first visible row after header
        
        'Highlight first 5 visible rows
        For Each cell In rng
            cnt = cnt + 1 'count for visible row
            If cnt = 5 Then
                Range(Cells(rCell, 1), Cells(cell.Row, lastColInt)).Select
                Selection.Interior.Color = vbYellow
                cnt = 0
                Exit For
            End If
        Next cell
        
    Next ky
    
    ActiveSheet.AutoFilterMode = False
    
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With

End Sub
