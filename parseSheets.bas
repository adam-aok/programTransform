Attribute VB_Name = "Module1"
'tested working 2021-01-26 1 PM WITHOUT parsefile function added
'sub to loop through files in a folder--the idea is to run the "parsefile" sub on each of these files, and output the data to the master revised sheet.
Sub LoopThroughFiles()
    Dim StrFile As String
    Dim SubFile As String
    Dim upperFolder As String
    Dim concatFolderName As String
    
    'set upperFolder to the master folder
    upperFolder = "\\d-peapcny.net\enterprise\P_Projects\83460\Design\Programming\"
    
    'set subFolder to the name of the directory you're looking in
    SubFolder = "DO NOT EDIT - 2020-12-15-840 SOA received from HA DEC 20201215\"
    
    
    'set strFile to any filename containing "SoA"
    StrFile = Dir(upperFolder + SubFolder + "*SoA*")
    
    Do While Len(StrFile) > 0
        Debug.Print StrFile
        parseFile (StrFile)
        StrFile = Dir
    Loop
End Sub

Sub parseFile(fName)

Dim wbFile As Workbook
Dim shX As Worksheet
Dim r As Range
Dim startPoint As Range
Dim endPoint As Range
Dim testR As Range

'need to pass entire filepath?
Workbooks(fName).Activate
Set wbFile = Workbooks(1)

'wbFile.Activate
'for each sheet in selected workbook
    For Each shX In wbFile.Sheets
        With shX.Cells
            Set startPoint = .Find("No. ", LookIn:=xlValues)
            If Not startPoint Is Nothing Then
                'MsgBox rngFound.Address
                'something is found
            Else
                'nothing is found
            End If
            
            Set endPoint = .Find("(I)", LookIn:=xlValues)
            If Not startPoint Is Nothing Then
                'MsgBox rngFound.Address
                'something is found
            Else
                'nothing is found
            End If
        End With
    testR = r
    'if found
    If Not r Is Nothing Then
        With shX
            For Each r In .Range(startPoint.Offset(1, 0), endPoint.End(xlDown).Offset(0, 2))
                r.Copy ActiveWorkbook.Worksheets("Sheet1").Range("A9")
            Next r
        End With
    End If
    Next shX
End Sub


Sub scrapeSoaMain()

Dim sourceDirectory As String
'set sourceDirectory to be the name of the place to search for files

parseFile (fName)
With Worksheets(soaFile).Cells
End With
End Sub
