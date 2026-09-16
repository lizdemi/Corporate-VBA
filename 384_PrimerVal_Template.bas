Attribute VB_Name = "Module1"
Public Platenum_Rng As Range, Platenum As Integer
Public Samplenum_Rng As Range, Samplenum As Integer
Public Repnum_Rng As Range, Repnum As Integer
Public Primernum_Rng As Range, Primernum As Integer
Public macroWb As Workbook, Template_A As Workbook, Template_B As Workbook
Public DataDims As Worksheet, Raw As Worksheet, PlateMap As Worksheet, Filter As Worksheet, Transposed As Worksheet
Public PlatemapRows As Variant, NamedCols As Variant
Public i As Integer, j As Integer, k As Integer

Sub Set_Variables()

Set macroWb = ThisWorkbook
Set DataDims = macroWb.ActiveSheet

Set Platenum_Rng = DataDims.Cells(4, 6)
Set Samplenum_Rng = DataDims.Cells(6, 6)
Set Repnum_Rng = DataDims.Cells(7, 6)
Set Primernum_Rng = DataDims.Cells(5, 6)

Platenum = Platenum_Rng.Value
Samplenum = Samplenum_Rng.Value
Repnum = Repnum_Rng.Value
Primernum = Primernum_Rng.Value

If Platenum = 0 Then
    Response = MsgBox("Number of plates will default to 1. Do you wish to proceed?", vbOKCancel)
    If Response = vbCancel Then
        Exit Sub
    Else:
        Platenum_Rng.Value = 1
    End If
ElseIf Samplenum = 0 Then
    Response = MsgBox("Maximum number of samples per primer will default to 8. Do you wish to proceed?", vbOKCancel)
    If Response = vbCancel Then
        Exit Sub
    Else:
        Samplenum_Rng.Value = 8
    End If
ElseIf Repnum = 0 Then
    Response = MsgBox("Maximum number of replicates per sample will default to 10. Do you wish to proceed?", vbOKCancel)
    If Response = vbCancel Then
        Exit Sub
    Else:
        Repnum_Rng.Value = 10
    End If
ElseIf Primernum = 0 Then
    MsgBox ("Please enter the total number of primers for this study")
    Exit Sub
End If

PlatemapRows = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P")
NamedCols = Array("Well Position", "Sample Name", "Sample", "Target", "Cq", "Cq Confidence", "Cq SD")


End Sub


Sub ChangeDimensions()

'Set public variables
Call Set_Variables

'Create the first workbook
Set Template_A = Workbooks.Add
DataDims.Range(DataDims.Cells(1, 1), DataDims.Cells(50, 24)).Copy
Template_A.ActiveSheet.Cells(1, 1).PasteSpecial
Template_A.ActiveSheet.Name = "Instructions"

macroWb.Activate
For i = 1 To Platenum
Sheets(Array("PlateMap_1", "Raw_1", "Filtered_1")).Select
Application.CutCopyMode = False
If i = 1 Then
    Sheets(Array("PlateMap_1", "Raw_1", "Filtered_1")).Copy _
        After:=Template_A.Sheets("Instructions")
    Template_A.Sheets("PlateMap_1").Name = "PlateMap_" & i
    Template_A.Sheets("Raw_1").Name = "Raw_" & i
    Template_A.Sheets("Filtered_1").Name = "Filtered_" & i
Else:
    Sheets(Array("PlateMap_1", "Raw_1", "Filtered_1")).Copy _
        After:=Template_A.Sheets("Filtered_" & i - 1)
    Template_A.Sheets("PlateMap_1 (2)").Name = "PlateMap_" & i
    Template_A.Sheets("Raw_1 (2)").Name = "Raw_" & i
    Template_A.Sheets("Filtered_1 (2)").Name = "Filtered_" & i
End If
'Set target dropdown
Template_A.Sheets("PlateMap_" & i).Unprotect
Template_A.Sheets("PlateMap_" & i).Cells(4, 1).Formula2 = "=UNIQUE(Raw_" & i & "!E:E,FALSE,FALSE)"
Template_A.Sheets("PlateMap_" & i).Protect
Template_A.Sheets("Filtered_" & i).Cells(1, 11).Validation.Delete
Template_A.Sheets("Filtered_" & i).Cells(1, 11).Validation.Add Type:=xlValidateList, _
        Formula1:="='PlateMap_" & i & "'!$A$4:$A$21"

Next i

If Samplenum <> 8 Or Repnum <> 10 Then
For i = 1 To Platenum
    Set Filter = Template_A.Sheets("Filtered_" & i)
    Range(Filter.Cells(3, 21), Filter.Cells(12, 31)).ClearFormats
    Range(Filter.Cells(3, 21), Filter.Cells(12, 31)).ClearContents
    Filter.Cells(3, 21).Formula2 = "=UNIQUE(L$3:L$390,FALSE,FALSE)"

    For j = 1 To Samplenum + 1
            Filter.Cells(3 + j, 22).Formula2 = "=TRANSPOSE(FILTER($N$3:$N$386,$L$3:$L$386=U" & 3 + j & "))"
    Next j
    
    For j = 1 To Repnum
        Filter.Cells(3, 21 + j).Value = j
    Next j

    Range(Filter.Cells(3, 21), Filter.Cells(4 + Samplenum, 21 + Repnum)).BorderAround Weight:=xlMedium

Next i
End If


Template_A.Sheets("Instructions").Activate
macroWb.Close

    

End Sub

Sub NewTemplate()
'For making a new document from cratch - time intensive!


'Set public variables
Call Set_Variables

'Create the first workbook
Set Template_A = Workbooks.Add
DataDims.Range(DataDims.Cells(1, 1), DataDims.Cells(50, 24)).Copy
Template_A.ActiveSheet.Cells(1, 1).PasteSpecial
Template_A.ActiveSheet.Name = "Instructions"

'Repeat for all plates
For i = 1 To Platenum

Call RawDataSheet(i)

Call PlatemapSheet(i)

Call FilterSheet(i)

Next i

End Sub

Sub RawDataSheet(plate As Integer)

'Create raw datasheet
If plate = 1 Then
    Template_A.Sheets.Add After:=Template_A.Sheets("Instructions")
Else:
    Template_A.Sheets.Add After:=Filter
End If
ActiveSheet.Name = "Raw_" & plate

Set Raw = Template_A.ActiveSheet

With Raw.Cells(1, 1)
    .Value = "qPCR data here"
    .Font.Bold = True
    .Font.Size = 14
    .BorderAround Weight:=xlMedium
End With
Columns("A").ColumnWidth = 23.86
End Sub

Sub PlatemapSheet(plate As Integer)

'Create platemap sheet
If plate = 1 Then
    Template_A.Sheets.Add After:=Template_A.Sheets("Instructions")
Else:
    Template_A.Sheets.Add After:=Filter
End If
ActiveSheet.Name = "PlateMap_" & plate

'Add the filter for later
Set PlateMap = Template_A.ActiveSheet
PlateMap.Cells(4, 1).Formula2 = "=UNIQUE(Raw_" & plate & "!E:E,FALSE,FALSE)"

'Create the platemap title border
PlateMap.Cells(3, 3).Value = "Paste your plate map here"
With Range(PlateMap.Cells(3, 3), PlateMap.Cells(3, 15))
    .Merge
    .Font.Bold = True
    .Font.Size = 18
    .Interior.Color = RGB(251, 226, 213)
    .BorderAround Weight:=xlMedium
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

'Add the platemap values
For j = 0 To 15
    PlateMap.Cells(6 + j, 3).Value = PlatemapRows(j)
Next j

For j = 1 To 24
    PlateMap.Cells(5, 3 + j).Value = j
Next j

With Range(PlateMap.Cells(5, 3), PlateMap.Cells(21, 27))
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlMedium
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

End Sub

Sub FilterSheet(plate As Integer)

'Create Filtered datasheet
Template_A.Sheets.Add After:=Raw
ActiveSheet.Name = "Filtered_" & plate

Set Filter = Template_A.ActiveSheet

'set assigned values
With Filter.Cells(1, 1)
    .Value = "Raw data organised by well number; Delete or insert columns as needed"
    .Font.Bold = True
    .Font.Size = 14
End With
Range(Filter.Cells(1, 1), Filter.Cells(1, 7)).Interior.Color = RGB(251, 226, 213)

Filter.Cells(2, 1).Value = NamedCols(0)
Filter.Cells(2, 2).Value = NamedCols(1)

With Filter.Cells(1, 10)
    .Value = "Target:"
    .Font.Bold = True
    .Font.Size = 11
    .Interior.Color = RGB(251, 226, 213)
End With

With Filter.Cells(1, 12)
    .Value = "Select Target from dropdown"
    .Font.Italic = True
    .Font.Color = RGB(255, 0, 0)
End With

Filter.Cells(1, 22).Value = "Copy into calculation template"
With Range(Filter.Cells(1, 22), Filter.Cells(1, 31))
    .Merge
    .Font.Italic = True
    .Font.Color = RGB(255, 0, 0)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(255, 255, 255)
    .BorderAround Weight:=xlMedium
End With

'Set filter table names
For j = 0 To 6
    Filter.Cells(2, 10 + j).Value = NamedCols(j)
Next j

'Set well number and sample name
c = 3
For j = 0 To 15
    For k = 1 To 24
        Filter.Cells(c, 1).Value = PlatemapRows(j) & k
        Filter.Cells(c, 2).Formula2 = "='" & PlateMap.Name & "'!" & PlateMap.Cells(6 + j, 3 + k).Address
        c = c + 1
    Next k
Next j

'Set xLookup formulas
For j = 2 To 386
    Filter.Cells(j, 3).Formula2 = "=XLOOKUP(TEXT($A" & j & ",0)," & PlateMap.Name & "!$B:$B," & PlateMap.Name & "!D:D,,2)"
    Filter.Cells(j, 4).Formula2 = "=XLOOKUP(TEXT($A" & j & ",0)," & PlateMap.Name & "!$B:$B," & PlateMap.Name & "!E:E,,2)"
    Filter.Cells(j, 5).Formula2 = "=XLOOKUP(TEXT($A" & j & ",0)," & PlateMap.Name & "!$B:$B," & PlateMap.Name & "!M:M,,2)"
    Filter.Cells(j, 6).Formula2 = "=XLOOKUP(TEXT($A" & j & ",0)," & PlateMap.Name & "!$B:$B," & PlateMap.Name & "!N:N,,2)"
    Filter.Cells(j, 7).Formula2 = "=XLOOKUP(TEXT($A" & j & ",0)," & PlateMap.Name & "!$B:$B," & PlateMap.Name & "!P:P,,2)"
Next j

'Set target dropdown
Filter.Cells(1, 11).Validation.Add Type:=xlValidateList, _
        Formula1:="='" & PlateMap.Name & "'!$A$4:$A$21"

'Set filter by target
Filter.Cells(3, 10).Formula2 = "=SORT(FILTER(A3:G386,ISNUMBER(SEARCH(K1,D3:D386))),3)"

'Create organised copies
Filter.Cells(3, 20).Formula2 = "=K1"
Filter.Cells(3, 21).Formula2 = "=UNIQUE(L$3:L$390,FALSE,FALSE)"


For j = 1 To Samplenum + 2
    Filter.Cells(3 + j, 22).Formula2 = "=TRANSPOSE(FILTER($N$3:$N$386,ISNUMBER(SEARCH(U" & 3 + j & ",$L$3:$L$386))))"
Next j

For j = 1 To Repnum
    Filter.Cells(3, 21 + j).Value = j
Next j


Range(Filter.Cells(3, 21), Filter.Cells(5 + Samplenum, 21 + Repnum)).BorderAround Weight:=xlMedium

End Sub
