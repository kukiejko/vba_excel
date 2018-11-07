Attribute VB_Name = "Excel_General"
Option Explicit
 
Sub ExcelDiet()
     
    Dim j               As Long
    Dim k               As Long
    Dim LastRow         As Long
    Dim LastCol         As Long
    Dim ColFormula      As Range
    Dim RowFormula      As Range
    Dim ColValue        As Range
    Dim RowValue        As Range
    Dim Shp             As Shape
    Dim ws              As Worksheet
     
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
    On Error Resume Next
     
    For Each ws In Worksheets
        With ws
             'Find the last used cell with a formula and value
             'Search by Columns and Rows
            On Error Resume Next
            Set ColFormula = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            Set ColValue = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            Set RowFormula = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            Set RowValue = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            On Error GoTo 0
             
             'Determine the last column
            If ColFormula Is Nothing Then
                LastCol = 0
            Else
                LastCol = ColFormula.Column
            End If
            If Not ColValue Is Nothing Then
                LastCol = Application.WorksheetFunction.Max(LastCol, ColValue.Column)
            End If
             
             'Determine the last row
            If RowFormula Is Nothing Then
                LastRow = 0
            Else
                LastRow = RowFormula.Row
            End If
            If Not RowValue Is Nothing Then
                LastRow = Application.WorksheetFunction.Max(LastRow, RowValue.Row)
            End If
             
             'Determine if any shapes are beyond the last row and last column
            For Each Shp In .Shapes
                j = 0
                k = 0
                On Error Resume Next
                j = Shp.TopLeftCell.Row
                k = Shp.TopLeftCell.Column
                On Error GoTo 0
                If j > 0 And k > 0 Then
                    Do Until .Cells(j, k).Top > Shp.Top + Shp.Height
                        j = j + 1
                    Loop
                    If j > LastRow Then
                        LastRow = j
                    End If
                    Do Until .Cells(j, k).Left > Shp.Left + Shp.Width
                        k = k + 1
                    Loop
                    If k > LastCol Then
                        LastCol = k
                    End If
                End If
            Next
             
            .Range(.Cells(1, LastCol + 1), .Cells(.Rows.Count, .Columns.Count)).EntireColumn.Delete
            .Range("A" & LastRow + 1 & ":A" & .Rows.Count).EntireRow.Delete
        End With
    Next
     
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
     
End Sub
 

