'Wrote this VBA script to aid building assessors in report writing.
'They are not too handy with spreadsheet and we wanted to avoid any mistakes when trying to copy a material type
'This work around lets the assessor paste his material type into a designated field
'and after pressing a button would automatically unfilter the sheet, insert a new row with all the relevant information, 
' and return the spreadsheet to the state it was originally in prior to hitting the macro.


Sub AddDupMatType()
    Dim w As Worksheet
    Dim filterArray()
    Dim currentFiltRange As String
    Dim col As Integer

    If Range("J2").Value = "" Then
            MsgBox "Please add a valid material type."
            Exit Sub
    End If
    
    Dim lookupval As String
    lookupval = Sheets("Asset Form").Range("J2")
    Dim rngfind As Range
    
    Set rngfind = Sheets("AllMatTypes").Range("A:A").Find(What:=lookupval, LookIn:=xlFormulas, MatchCase:=False)
    
    If rngfind Is Nothing Then
        MsgBox "Please add a valid material type."
        Exit Sub
    End If
    
    Set w = Sheets("Asset Form")
    ' Capture AutoFilter settings
    With w.AutoFilter
        currentFiltRange = .Range.Address
        With .Filters
            ReDim filterArray(1 To .Count, 1 To 3)
            For f = 1 To .Count
                With .Item(f)
                    If .On Then
                        filterArray(f, 1) = .Criteria1
                        If .Operator Then
                            filterArray(f, 2) = .Operator
                            'filterArray(f, 3) = .Criteria2 'simply delete this line to make it work in Excel 2010
                        End If
                    End If
                End With
            Next f
        End With
    End With

    'Remove AutoFilter
    w.AutoFilterMode = False

'____________________________________


    Set rngfind = Sheets("Asset Form").Range("J4:J1000").Find(What:=lookupval, LookIn:=xlFormulas, MatchCase:=False)
    
    If rngfind Is Nothing Then
            MsgBox "Please add a valid material type."
            GoTo Restore_Filter
    End If
    
    

        Range("J4:J1000").Select
        Selection.Find(What:=lookupval, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        Rows(ActiveCell.Row).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown


'__________________________________________
Restore_Filter:
    ' Restore Filter settings
    For col = 1 To UBound(filterArray(), 1)
        If Not IsEmpty(filterArray(col, 1)) Then
            If filterArray(col, 2) Then
                w.Range(currentFiltRange).AutoFilter field:=col, _
                Criteria1:=filterArray(col, 1), _
                Operator:=filterArray(col, 2), _
                Criteria2:=filterArray(col, 3)
            Else
                w.Range(currentFiltRange).AutoFilter field:=col, _
                Criteria1:=filterArray(col, 1)
            End If
        End If
    Next col
    If w.AutoFilterMode = False Then
        Range("A3:AS1000").Select
        Selection.AutoFilter
    End If
    
End Sub



