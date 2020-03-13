VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'L and populations has to be "public" variables as they are being used by multible functions
Public L As Long
Public my_path As String
Public populations As Integer


Sub YesNo()
    
Call SelectEnvironments

End Sub

Sub SelectEnvironments()
    
    'Deleting all sheets except first
    Dim sht As Worksheet
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name <> "S1" Then
            sht.Delete
        End If
    Next
    
    Application.DisplayAlerts = True
    
    'Clearing contents of first sheet
    Sheets("S1").UsedRange.ClearContents
    
    'Writing some loops that results in the same as copying template but without needing
    'to refer to another file
    With Sheets("S1")
    .Cells(1, "A") = "A"
    .Cells(1, "B") = "Environment"
    .Cells(1, "D") = "Falcon 50ml"
    .Cells(1, "E") = 1
    .Cells(1, "G") = 800
    'Hardcoding sample line 2 which values will be replaced by later lines of code
    .Cells(2, "A") = "D"
    .Cells(2, "B") = "Plate1"
    .Cells(2, "D") = "96 Well Microplate"
    .Cells(2, "E") = 10
    .Cells(2, "G") = 200
    .Cells(2, "K") = "96 Well Microplate"
    q1 = 11
    For q2 = 3 To 61
        .Cells(q2, "A") = "D"
        .Cells(q2, "B") = "=B2"
        .Cells(q2, "D") = "96 Well Microplate"
        .Cells(q2, "E") = q1
        .Cells(q2, "G") = "=G2"
        .Cells(q2, "K") = "96 Well Microplate"
        q1 = q1 + 1
    Next q2
    .Cells(62, "A") = "W"
    End With

    'Defining variables
    Dim ws As Worksheet
    Dim x As Integer
    Dim env As Variant
    Dim volume As Integer
    Dim y As Double
    
    'Asking for inputs
    env = InputBox("Number of Environments:")
    volume = InputBox("Volume (ul) of environment in each well:")
    Sheets("S1").Range("G2").Value = volume
    
    'Filling in values and inserting new rows
    y = 960 / volume
    y = Application.WorksheetFunction.RoundDown(y, 0)
    Sheets("S1").Range("G1").Value = volume * y

    Dim R As Integer
    Dim w As Integer
    Dim Ax As Integer
    Ax = 60 / y - 1
    w = 0
    For R = 1 To Ax
        Range("A" & y + 2 + w).EntireRow.Insert
        Range("A1:G1").Copy _
        Destination:=Range("A" & y + 2 + w)
        Cells(y + 2 + w, "E") = Cells(1, "E").Value
        w = w + y + 1
    Next R
    
    w = 0
    
    'Copying environment sheets
    For x = 1 To env - 1
        Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook _
        .Sheets("S" & x))
        ws.Name = "S" & x + 1
        Sheets("S1").Range("A1:K100").Copy _
        Destination:=Sheets("S" & x + 1).Range("A1")
        Range("B2").Value = "Plate" & x + 1
        Range("E1").Value = x + 1
        
        ' Filling out values in secondary sheets
        For R = 1 To Ax
            Cells(y + 2 + w, "E") = Cells(1, "E").Value
            w = w + y + 1
        Next R
        w = 0
    Next x
    
End Sub

Public Sub pops()

    'Generating one worksheet for each plate
    Dim plate As Worksheet
    L = Sheets.Count
    For b = 0 To L - 1
        Set plate = Sheets.Add(After:=Sheets.Item(L))
        plate.Name = "Plate" & L - b
    Next b
    
    'Asking for population volume
    populations = InputBox("Number of populations:")
    Dim popvol As Integer
    popvol = InputBox("Volume (ul) of population in each well:")
    
    'Filling in values for population sheets
    For c = 1 To L
    w = 1
    With Sheets("Plate" & c)
    For b = 1 To 6 * populations - 5 Step 6
        .Cells(b, "A") = "A"
        .Cells(b, "B") = "Population"
        .Cells(b, "D") = "falcon 15ml"
        .Cells(b, "E") = (b - 1) / 6 + 1
        .Cells(b, "G") = 4 * popvol
        .Cells(b + 5, "A") = "W"
        For k = b + 1 To b + 4
            .Cells(k, "A") = "D"
            .Cells(k, "B") = "Plate" & c
            .Cells(k, "D") = "96 Well Microplate"
            .Cells(k, "G") = popvol
            .Cells(w, "I") = WorksheetFunction.RandBetween(10000, 100000) / 100000
            w = w + 1
        Next k
    Next b
    
    'Ranking randomly generated numbers
    Sheets("Plate" & c).Activate
    u = 1
    For b = 1 To 6 * populations - 5 Step 6
        For k = b + 1 To b + 4
            .Cells(u, "J") = Application.WorksheetFunction.Rank_Eq(Range("I" & u), Range("I1:I60"), 1)
            u = u + 1
        Next k
    Next b
    'Making extra column with actual well numbers in K-column
    g = 1
    'g is the row number
        
    For x = 10 To 82 Step 8
    'Step is the number of wells to jump up between consecutive operaitons of the loop.
    
        For y = x To x + 5
        'y is the value in a cell
        'y = x is the well number
        
            .Cells(g, "K") = y
            'function that is setting row g, column "K" = y
            
            g = g + 1
            'g increase the well by 1 until x + 5 before stepping with 8
                        
        Next y
    Next x
    
    'Filling in ranked numbers from J-column to E-column
    w = 1
    For b = 1 To 6 * populations - 5 Step 6
        For k = b + 1 To b + 4
            .Cells(k, "E") = Range("J" & w)
            w = w + 1
        Next k
    Next b
    
    'Relating number in E-column to well number from J-column
    For b = 1 To 6 * populations - 5 Step 6
        For k = b + 1 To b + 4
            x = .Cells(k, "E")
            .Cells(k, "E") = Range("K" & x)
        Next k
    Next b
    
    'Deleting intermediate results
    Range("I1:K60").Delete
    
    End With
    Next c
    'Sending program to the extract function where data is copied to the overview-file
    Call Water
    
End Sub

Sub Water()
    
    Set wa = Sheets.Add(before:=Sheets(1))
    wa.Name = "Water"
    
        With Sheets("Water")
        w = 1
        For x = 1 To L
            For q = 1 To 12
                .Cells(w, "A") = "A"
                .Cells(w, "B") = "Water"
                .Cells(w, "D") = "Falcon 50ml"
                .Cells(w, "E") = 1
                .Cells(w, "G") = 750
            w = w + 4
            Next q
        Next x
        
        Dim j As Integer
        Dim i As Integer
        Dim g As Integer
        
        j = 2
        For i = 1 To L
        g = 1
            For a = 1 To 12
                For b = 1 To 3
                .Cells(j, "A") = "D"
                .Cells(j, "B") = "Plate" & i
                .Cells(j, "D") = "96 Well Microplate"
                .Cells(j, "E") = g
                .Cells(j, "G") = 250
                j = j + 1
                If a <= 8 Then g = g + 8
                If a >= 9 Then g = g + 1
                Next b
            If a = 4 Then g = 8
            If a = 8 Then g = 2
            If a = 10 Then g = 90
            j = j + 1
            Next a
        Next i
        
        End With
        
    Call Extract
    
End Sub

Public Sub Extract()
    
    Set ov = Sheets.Add(After:=ActiveWorkbook.Sheets("Plate" & L))
    ov.Name = "overview"
    
    'Copying well number to overview
    Dim ar As Variant
    ar = VBA.Array(1, 2, 3, 4)
    ar = Application.WorksheetFunction.Transpose(ar)
    Dim q As Integer
    q = 2
    For c = 1 To L
    
        For b = 2 To 6 * populations - 4 Step 6
            Sheets("Plate" & c).Range(Cell1:="E" & b, Cell2:="E" & b + 3).Copy
            Sheets("overview").Range(Cell1:="F" & q, Cell2:="F" & q + 3).PasteSpecial Paste:=xlPasteValues
            Sheets("overview").Range(Cell1:="C" & q, Cell2:="C" & q + 3).Value = ar
            Sheets("Plate" & c).Range("E" & b - 1).Copy
            Sheets("overview").Range(Cell1:="B" & q, Cell2:="B" & q + 3).PasteSpecial Paste:=xlPasteValues
            q = q + 4
        Next b
        
        t = c * 60 - 58
        Sheets("overview").Range(Cell1:="A" & t, Cell2:="A" & t + 59).Value = c
        Sheets("overview").Range(Cell1:="D" & t, Cell2:="D" & t + 59).Value = c
        
    Next c
    Call look_up_well
    
End Sub

Sub look_up_well()
    w = 1
    Dim la As Variant
    For x = 2 To 11
        la = VBA.Array("B" & x, "C" & x, "D" & x, "E" & x, "F" & x, "G" & x)
        la = Application.WorksheetFunction.Transpose(la)
        Sheets("overview").Range(Cell1:="J" & w, Cell2:="J" & w + 5) = la
        w = w + 6
    Next x
    
    With Sheets("overview")
    g = 1
    For x = 10 To 82 Step 8
        For y = x To x + 5
            .Cells(g, "I") = y
            g = g + 1
        Next y
    Next x
    
    Dim wellnr
    Dim pos
    
    For x = 2 To 60 * L + 1
        wellnr = Range("F" & x).Value
        pos = Application.Match(wellnr, Range("I1:I60"), 0)
        Range("E" & x).Value = Range("J" & pos).Value
    Next x
    
    Range("I1:J60").Delete
    End With
    Call headlines
    
End Sub

Sub headlines()
    
    Sheets("overview").Cells(1, "A") = "Environment"
    Sheets("overview").Cells(1, "B") = "Population"
    Sheets("overview").Cells(1, "C") = "Replicate"
    Sheets("overview").Cells(1, "D") = "Plate"
    Sheets("overview").Cells(1, "E") = "Well ID"
    Sheets("overview").Cells(1, "F") = "Well #"
    
End Sub

Sub Save_overview()
    
    my_path = ActiveWorkbook.Path
    Sheets("overview").Move
    ActiveWorkbook.SaveAs Filename:=my_path & "/" & "overview" & ".xlsx", FileFormat:=xlWorkbookDefault
    
    Call SaveToGWL

End Sub

Sub SaveToGWL()
    
    
    ' Creating new workbook of sheet one
    ThisWorkbook.Sheets("Water").Copy
    ActiveWorkbook.SaveAs Filename:=my_path & "/" & "Script" & ".csv", _
    FileFormat:=xlCSV, Local:=True
    
    Dim wbs As Workbook
    Set wbs = Workbooks("Script.csv")
    
    ' Copying all tables to sheet one of new workbook
    Dim ws As Worksheet
    Dim wssource As Worksheet
    Set wssource = ThisWorkbook.Sheets("Water")
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> wssource.Name Then
            ws.Range("A1:K150").Copy
            wbs.Sheets("Script").Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        End If
    Next ws
    
    ' Save as gwl file
    With wbs
        .SaveAs my_path & "/" & "Script" & ".gwl", FileFormat:=xlCSV, Local:=True
        .Close False
    End With
    Kill my_path & "/" & "Script.csv"
End Sub

