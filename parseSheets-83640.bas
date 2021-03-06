'tested working 2021-01-26 1 PM WITHOUT parsefile function added
'sub to loop through files in a folder--the idea is to run the "parsefile" sub on each of these files, and output the data to the master revised sheet.
Sub LoopThroughFiles()
    Dim StrFile As String
    Dim SubFile As String
    Dim upperFolder As String
    Dim concatFolderName As String
    
    'set upperFolder to the master folder
    upperFolder = "\\d-peapcny.net\enterprise\P_Projects\83460\Design\Programming\UDP_Program-DataMining_83460.00.0\"
    
    'set subFolder to the name of the directory you're looking in
    SubFolder = "1_CLIENTFORMAT\"
    
    
    'set strFile to any filename containing "SoA"
    StrFile = Dir(upperFolder + SubFolder + "*SoA*")
    
    'add strFile to complete folder path
    concatFolderName = upperFolder + SubFolder + StrFile
    
    Do While Len(StrFile) > 0
        Debug.Print StrFile
        parseFile (StrFile)
        StrFile = Dir
    Loop
End Sub

Sub parseFile(fName)

'Sub parseFile()

Dim wbFile As Workbook
Dim shX As Worksheet
Dim r As Range
Dim startPoint As Range
Dim startHeader As Range
Dim startBottom As Range

Dim endPoint As Range
Dim testR As Range
Dim foundRange As Range
Dim dirPath As String

Dim columnDepth As Long

Dim destStart As Range
Dim soaCols() As Variant
Dim destCols() As Variant
Dim wsName As String
Dim sourceRange, fillRange As Range
Dim newOrExist As String
Dim deptName As String

soaCols = Array("(A)", "(B)", "(C)", "(D)", "(E)", "(F)", "(G)", "(H)")
destCols = Array("SoA Ref No.", "(B)", "(C)", "Quantity of Rooms", "NOFA (m2)", "(F)", "(G)", "Remarks")

dirPath = "\\d-peapcny.net\enterprise\P_Projects\83460\Design\Programming\UDP_Program-DataMining_83460.00.0\1_CLIENTFORMAT\"

'fName = "017-20200630_PnC_LKB SoA Section 17 Dangerous Good Stores.xlsx"

'need to pass entire filepath?
Workbooks.Open Filename:=dirPath + fName
'Set wbFile = Workbooks(dirPath & fName)

Set wbFile = Workbooks(fName)
'for each sheet in selected workbook
    For Each shX In wbFile.Sheets
        
        'find the start of each column, and descend down the column to find the bottom
        With shX.Cells
        
            'shX.Cells.UnMerge
            Set startHeader = .Find("(A)", LookIn:=xlValues)
            
            'if the startPoint is found, then descend down the column to find the bottom-left value
            If Not startHeader Is Nothing Then
                'MsgBox rngFound.Address
                'startBottom = .Find(What:="Note 1: "
                'Set startBottom = .Find("Note 1: ", LookIn:=xlValues)
                'something is found
                'Set foundRange = .Range(startHeader.Offset(4, 0), startBottom.Offset(-3, 0))
            Else
                'nothing is found
                
            End If
            
            'find the far-right column
            Set startBottom = .Find("Note 1: ", LookIn:=xlValues)
            
            'if the end column header is found
            If Not startBottom Is Nothing Then
                'MsgBox rngFound.Address
                'something is found
                '.Range
                Set foundRange = .Range(startHeader.Offset(4, 0), startBottom.Offset(-5, 11))
                'foundRange.UnMerge
                'foundRange.Columns("C:E").Delete
                'foundRange.Columns(4).Delete
                'foundRange.Columns(1).Insert , xlShiftToRight
                
                'Set sourceRange = .Fin
                'Set foundRange.Range(Cells(1, 1)).Value = fName
                'foundRange.Range(Cells(1, 1)).AutoFill Destination:=foundRange.Columns(1)
                'foundRange(foundRange.Rows.Count
                'if
            Else
                'nothing is found
            End If
            
            If Not .Find("Hospital Authority", LookIn:=xlValues) Is Nothing Then
                deptName = .Find("Hospital Authority", LookIn:=xlValues).Offset(1, 0).Value
            End If
            
                        
            'determine new or existing
            If Not .Find("(New Block)", LookIn:=xlValues) Is Nothing Then
                newOrExist = "New"
            Else
                newOrExist = "Existing"
            End If
            
            'now (if the range has been found, pasting each set of rows from this sheet into the master workbook sheet
            
            
        End With
  
        
        'find the start of each column, and descend down the column to find the bottom
        With Workbooks("OUTPUT_2021-04-23_LKB.xlsm").Worksheets("Data").Cells
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
                .UnMerge
                foundRange.Columns("A:B").Copy Destination:=destStart
                foundRange.Columns("F:L").Copy Destination:=destStart.Offset(0, 2)
                .Range(destStart.Offset(0, -5), destStart.Offset(foundRange.Rows.Count, -5)).Value = fName
                .Range(destStart.Offset(0, -4), destStart.Offset(foundRange.Rows.Count, -4)).Value = shX.Name
                .Range(destStart.Offset(0, -3), destStart.Offset(foundRange.Rows.Count, -3)).Value = deptName
                .Range(destStart.Offset(0, 11), destStart.Offset(foundRange.Rows.Count, 9)).Value = newOrExist
                
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


