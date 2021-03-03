'Sub parseFile(fName)

Sub parseFile()

Dim wbFile As Workbook
Dim shX As Worksheet
Dim R As Range
Dim startPoint As Range
Dim startHeader As Range
Dim startBottom As Range

Dim endPoint As Range
Dim testR As Range
Dim foundRange As Range
Dim dirPath As String
'Dim fName As String

Dim columnDepth As Long

Dim destStart As Range
Dim soaCols() As Variant
Dim destCols() As Variant
Dim wsName As String
Dim sourceRange, fillRange As Range
Dim newOrExist As String
Dim deptName As String
Dim x As Integer

Dim containerRange As Integer

soaCols = Array("(A)", "(B)", "(C)", "(D)", "(E)", "(F)", "(G)", "(H)")
destCols = Array("SoA Ref No.", "(B)", "(C)", "Quantity of Rooms", "NOFA (m2)", "(F)", "(G)", "Remarks")

'dirPath = "\\d-peapcny.net\enterprise\P_Projects\83460\Design\Programming\DO NOT EDIT - 2020-12-15-840 SOA received from HA DEC 20201215\"

fName = "02.16.21_Santiago Hospital_Space program - Translated with client edits.xlsx"

'fName = "009-20200630_PnC_LKB SoA Section 9 Information Counter.xlsx"

'need to pass entire filepath?
'Workbooks.Open Filename:=dirPath + fName
'Set wbFile = Workbooks(dirPath & fName)

Set wbFile = Workbooks(fName)
'for each sheet in selected workbook
    For Each shX In wbFile.Sheets
        If shX.Name <> "SUMMARY" And shX.Name <> "Colors" And shX.Name <> "BASE RECEIVED" Then
        
        'find the start of each column, and descend down the column to find the bottom
        With shX.Cells
        
            'shX.Cells.UnMerge
            'find start point for the top-left limits of the copy selection
            Set startHeader = .Find("Programa Funcional - HOSPITAL SANTIAGO", LookIn:=xlValues)
            
            'if the startPoint is found, then descend down the column to find the bottom-left value
            If Not startHeader Is Nothing Then
                
            Else
                'nothing is found
                
            End If
            
            'find the far-right column
            Set startBottom = .Range("B1048576").End(xlUp)
            
            If Not startBottom Is Nothing Then
                MsgBox startBottom.Address
                Set foundRange = .Range(startHeader.Offset(10, 0), startBottom.Offset(0, 6))
                'For Each r In .Range(startHeader.Offset(10, 0), startBottom.Offset(0, 6)).Rows
                    'If r.Cells(0, 1) <> "" Then
                        'If foundRange Is Nothing Then
                            'Set foundRange = r
                        'Else
                            'Set foundRange = Union(foundRange, r)
                        'End If
                    'End If
                'Next r
            Else
                'nothing is found
            End If
            
            'now (if the range has been found, pasting each set of rows from this sheet into the master workbook sheet
            
        End With
        End If
        
        'now (if the range has been found, pasting each set of rows from this sheet into the master workbook sheet
        'find the start of each column, and descend down the column to find the bottom
        With Workbooks("Data-Intake_AOK.xlsm").Worksheets("Sheet1").Cells
            If Not foundRange Is Nothing Then
                If shX.Name <> "Guidelines" Then
                Set destStart = .Range("F1048576").End(xlUp).Offset(1, 0)
                'MsgBox destStart.Address
                'get subDepartment
                
                'For Each r In foundRange.Rows
                   ' If Len(r("A1").Value) = 3 Then
                  '      Set destStart.Offset(0, -1) = r("A1")
                 '   End If
                'Next r
                    
                'For Each r In foundRange.Rows
                     'If r.Cells("I1").Value <> "" Then
                        
                  '      Set destStart.Offset(0, -1) = Cellular
                 '   End If
                'Next r
                'Workbooks("2021-02-19_LKB Proposed Excel Format Headings.xlsm").Worksheets("Sheet1").Cells.UnMerge
                foundRange.Columns("A:I").Copy
                destStart.PasteSpecial Paste:=xlPasteValues
                'foundRange.Columns("F:L").Copy Destination:=destStart.Offset(0, 2)
                .Range(destStart.Offset(0, -5), destStart.Offset(foundRange.Rows.Count, -5)).Value = fName
                .Range(destStart.Offset(0, -4), destStart.Offset(foundRange.Rows.Count, -4)).Value = shX.Name
                .Range(destStart.Offset(0, -3), destStart.Offset(foundRange.Rows.Count, -3)).Value = deptName
                .Range(destStart.Offset(0, 9), destStart.Offset(foundRange.Rows.Count, 9)).Value = newOrExist
                
                'If
                'destStart.FillDown
                'For Each r In foundRange.Rows
                '    Set destStart = Workbooks("2021-02-17_LKB Proposed Excel Format Headings.xlsm").Worksheets("Sheet1").Range("I1048576").End(xlUp)
                '    MsgBox destStart.Address
                '   'r.Copy Destination:=destStart
                'Next r
                
                    
                End If
            End If
        End With
    'testR = r
    'if found
    'foundRange = shX.Range(startPoint.Offset(10, 0), endPoint.End(xlDown).Offset(0, 2))

    
    Next shX
    
Workbooks(fName).Close SaveChanges:=False
End Sub
