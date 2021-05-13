'Derivative of 83640, but adapted to differing format

'sub to loop through files in a folder--the idea is to run the "parsefile" sub on each of these files, and output the data to the master revised sheet.
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

fName = "02.16.21_Santiago Hospital_Space program - Translated with client edits - copy.xlsx"

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
                'MsgBox startBottom.Address
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
        With Workbooks("OUTPUT_86050.00.0_Data-Intake_AOK.xlsm").Worksheets("RawData").Cells
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


'Sub scrapeSoaMain()

'Dim sourceDirectory As String
'set sourceDirectory to be the name of the place to search for files

'parseFile (fName)
'With Worksheets(soaFile).Cells
'End With
'End Sub


Sub ColorRGB(Rs As Range, Rc As Range)
Dim R As Long
Dim G As Long
Dim B As Long
Dim Address(1 To 3) As Long
Dim I As Integer: I = 1
For Each Cell In Rs.Cells
Address(I) = Cell.Value
I = I + 1
Next
R = Address(1)
G = Address(2)
B = Address(3)
Rc.Interior.Color = RGB(R, G, B)
End Sub

Public Function getColor(rng As Range, ByVal ColorFormat As String) As Variant
    Dim ColorValue As Variant
    ColorValue = Cells(rng.Row, rng.Column).Interior.Color
    Select Case LCase(ColorFormat)
        Case "index"
            getColor = rng.Interior.ColorIndex
        Case "rgb"
            getColor = (ColorValue Mod 256) & ", " & ((ColorValue \ 256) Mod 256) & ", " & (ColorValue \ 65536)
        Case Else
            getColor = "Only use 'Index' or 'RGB' as second argument!"
    End Select
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function            Color
'   Purpose             Determine the Background Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Color(rng As Range, Optional formatType As Integer = 0) As Variant
    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Interior.Color
    Select Case formatType
        Case 1
            Color = Hex(colorVal)
        Case 2
            Color = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            Color = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            Color = colorVal
    End Select
End Function

Public Function getColorAOK(rng As Range) As Long

    getColorAOK = rng.Cell.Interior.Color
End Function