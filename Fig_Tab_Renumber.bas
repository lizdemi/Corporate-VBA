Attribute VB_Name = "Module1"
Option Explicit

Public sld As Slide, shp As Shape
Public startSlideIndex As Long
Public StartingFigure As Boolean, StartingTable As Boolean
Public currentText As String
Public prev_figureNumber As Long, figureNumber As Long
Public prev_TableNumber As Long, TableNumber As Long
Public colonPosition As Long
Public figurePrefix As String, TablePrefix As String
Public remainingText As String

Sub RenumberFigures()
'Ensure a slide is selected
If ActiveWindow.Selection.Type <> ppSelectionSlides Then
    MsgBox "Please select the slide where renumbering should begin.", vbExclamation
    Exit Sub
End If

'Set the selected slide as startSlideIndex value
startSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex

'Clear a previous run
StartingFigure = False


'Run through all slides from the selected slide onwards, checking all shapes for text
For Each sld In ActivePresentation.Slides
    If sld.SlideIndex >= startSlideIndex Then
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    currentText = Trim(shp.TextFrame.TextRange.Text) 'Contain the text as a string variable
                    
                    'Now that all text is saved as a string, check if the string itself completes all syntax requirements
                    If Left(currentText, 7) = "Figure " Then 'Check that text starts with "Figure "
                        colonPosition = InStr(currentText, ":")
                        If colonPosition > 8 Then 'Ensure a ':' exists after the figure number
                            figurePrefix = Mid(currentText, 8, colonPosition - 8) 'extract the number
                            If IsNumeric(Trim(figurePrefix)) Then 'Verify the extracted value is numeric
                                figureNumber = CLng(Trim(figurePrefix)) 'save number as variable
                                remainingText = Mid(currentText, colonPosition) 'save remaining text as a String
                                
                                'Operations will now vary between the first slide and the subsequent slides.
                                If sld.SlideIndex = startSlideIndex And StartingFigure = False Then 'For the first figure on the first slide only
                                    StartingFigure = True
                                    prev_figureNumber = figureNumber
                                'Print error if the slide does not contain a figure value
                                ElseIf StartingFigure = False Then
                                        MsgBox "No Figure caption found on the selected slide.", vbExclamation
                                        Exit Sub
                                'Still on the first slide, if there are multiple figures on the slide that match the syntax
                                ElseIf sld.SlideIndex = startSlideIndex Then 'For the second figure on the first slide
                                    figureNumber = prev_figureNumber + 1
                                    shp.TextFrame.TextRange.Text = "Figure " & figureNumber & remainingText
                                    prev_figureNumber = figureNumber
                                
                                'For all remaining slides (because >= startSlideIndex was already established and we have now excluded =)
                                Else:
                                    'Increase the value
                                    figureNumber = prev_figureNumber + 1
                                    shp.TextFrame.TextRange.Text = "Figure " & figureNumber & remainingText
                                    prev_figureNumber = figureNumber
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next shp
    End If
Next sld


End Sub

Sub RenumberTables()
'Ensure a slide is selected
If ActiveWindow.Selection.Type <> ppSelectionSlides Then
    MsgBox "Please select the slide where renumbering should begin.", vbExclamation
    Exit Sub
End If

'Set the selected slide as startSlideIndex value
startSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex

'Clear a previous run
StartingTable = False

'Run through all slides from the selected slide onwards, checking all shapes for text
For Each sld In ActivePresentation.Slides
    If sld.SlideIndex >= startSlideIndex Then
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    currentText = Trim(shp.TextFrame.TextRange.Text) 'Contain the text as a string variable
                    
                    'Now that all text is saved as a string, check if the string itself completes all syntax requirements
                    If Left(currentText, 6) = "Table " Then 'Check that text starts with "Table "
                        colonPosition = InStr(currentText, ":")
                        If colonPosition > 7 Then 'Ensure a ':' exists after the Table number
                            TablePrefix = Mid(currentText, 7, colonPosition - 7) 'extract the number
                            If IsNumeric(Trim(TablePrefix)) Then 'Verify the extracted value is numeric
                                TableNumber = CLng(Trim(TablePrefix)) 'save number as variable
                                remainingText = Mid(currentText, colonPosition) 'save remaining text as a String
                                
                                'Operations will now vary between the first slide and the subsequent slides.
                                If sld.SlideIndex = startSlideIndex And StartingTable = False Then 'For the first Table on the first slide only
                                    StartingTable = True
                                    prev_TableNumber = TableNumber
                                    'Print error if the slide does not contain a Table value
                                ElseIf StartingTable = False Then
                                        MsgBox "No Table caption found on the selected slide.", vbExclamation
                                        Exit Sub
                                'Still on the first slide, if there are multiple Tables on the slide that match the syntax
                                ElseIf sld.SlideIndex = startSlideIndex Then 'For the second Table on the first slide
                                    TableNumber = prev_TableNumber + 1
                                    shp.TextFrame.TextRange.Text = "Table " & TableNumber & remainingText
                                    prev_TableNumber = TableNumber
                                
                                'For all remaining slides (because >= startSlideIndex was already established and we have now excluded =)
                                Else:
                                    'Increase the value
                                    TableNumber = prev_TableNumber + 1
                                    shp.TextFrame.TextRange.Text = "Table " & TableNumber & remainingText
                                    prev_TableNumber = TableNumber
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next shp
    End If
Next sld

StartingTable = Empty


End Sub

Sub RenumberSupplementalFigures()
'Ensure a slide is selected
If ActiveWindow.Selection.Type <> ppSelectionSlides Then
    MsgBox "Please select the slide where renumbering should begin.", vbExclamation
    Exit Sub
End If

'Set the selected slide as startSlideIndex value
startSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex

'Clear a previous run
StartingFigure = False


'Run through all slides from the selected slide onwards, checking all shapes for text
For Each sld In ActivePresentation.Slides
    If sld.SlideIndex >= startSlideIndex Then
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    currentText = Trim(shp.TextFrame.TextRange.Text) 'Contain the text as a string variable
                    
                    'Now that all text is saved as a string, check if the string itself completes all syntax requirements
                    'Try Figure SX syntax first
                    If Left(currentText, 8) = "Figure S" Then 'Check that text starts with "Figure "
                        colonPosition = InStr(currentText, ":")
                        If colonPosition > 9 Then 'Ensure a ':' exists after the figure number
                            figurePrefix = Mid(currentText, 9, colonPosition - 9) 'extract the number
                            If IsNumeric(Trim(figurePrefix)) Then 'Verify the extracted value is numeric
                                figureNumber = CLng(Trim(figurePrefix)) 'save number as variable
                                remainingText = Mid(currentText, colonPosition) 'save remaining text as a String
                                
                                'Operations will now vary between the first slide and the subsequent slides.
                                If sld.SlideIndex = startSlideIndex And StartingFigure = False Then 'For the first figure on the first slide only
                                    StartingFigure = True
                                    prev_figureNumber = figureNumber
                                'Print error if the slide does not contain a figure value
                                ElseIf StartingFigure = False Then
                                        MsgBox "No Figure caption found on the selected slide.", vbExclamation
                                        Exit Sub
                                'Still on the first slide, if there are multiple figures on the slide that match the syntax
                                ElseIf sld.SlideIndex = startSlideIndex Then 'For the second figure on the first slide
                                    figureNumber = prev_figureNumber + 1
                                    shp.TextFrame.TextRange.Text = "Figure S" & figureNumber & remainingText
                                    prev_figureNumber = figureNumber
                                
                                'For all remaining slides (because >= startSlideIndex was already established and we have now excluded =)
                                Else:
                                    'Increase the value
                                    figureNumber = prev_figureNumber + 1
                                    shp.TextFrame.TextRange.Text = "Figure S" & figureNumber & remainingText
                                    prev_figureNumber = figureNumber
                                End If
                            End If
                        End If
                    
                    'Try Supplemental Figure X syntax next
                    ElseIf Left(currentText, 20) = "Supplemental Figure " Then 'Check that text starts with "Figure "
                        colonPosition = InStr(currentText, ":")
                        If colonPosition > 21 Then 'Ensure a ':' exists after the figure number
                            figurePrefix = Mid(currentText, 21, colonPosition - 21) 'extract the number
                            If IsNumeric(Trim(figurePrefix)) Then 'Verify the extracted value is numeric
                                figureNumber = CLng(Trim(figurePrefix)) 'save number as variable
                                remainingText = Mid(currentText, colonPosition) 'save remaining text as a String
                                
                                'Operations will now vary between the first slide and the subsequent slides.
                                If sld.SlideIndex = startSlideIndex And IsEmpty(StartingFigure) Then 'For the first figure on the first slide only
                                    StartingFigure = figureNumber
                                    prev_figureNumber = StartingFigure
                                    'Print error if the slide does not contain a figure value
                                    If IsEmpty(StartingFigure) Then
                                            MsgBox "No Figure caption found on the selected slide.", vbExclamation
                                            Exit Sub
                                    End If
                                'Still on the first slide, if there are multiple figures on the slide that match the syntax
                                ElseIf sld.SlideIndex = startSlideIndex Then 'For the second figure on the first slide
                                    figureNumber = prev_figureNumber + 1
                                    shp.TextFrame.TextRange.Text = "Supplemental Figure " & figureNumber & remainingText
                                    prev_figureNumber = figureNumber
                                
                                'For all remaining slides (because >= startSlideIndex was already established and we have now excluded =)
                                Else:
                                    'Increase the value
                                    figureNumber = prev_figureNumber + 1
                                    shp.TextFrame.TextRange.Text = "Supplemental Figure " & figureNumber & remainingText
                                    prev_figureNumber = figureNumber
                                End If
                            End If
                        End If
                    
                    End If
                End If
            End If
        Next shp
    End If
Next sld
    
StartingFigure = Empty

End Sub

Sub RenumberSupplementalTables()
'Ensure a slide is selected
If ActiveWindow.Selection.Type <> ppSelectionSlides Then
    MsgBox "Please select the slide where renumbering should begin.", vbExclamation
    Exit Sub
End If

'Set the selected slide as startSlideIndex value
startSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex

'Clear a previous run
StartingTable = False

'Run through all slides from the selected slide onwards, checking all shapes for text
For Each sld In ActivePresentation.Slides
    If sld.SlideIndex >= startSlideIndex Then
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    currentText = Trim(shp.TextFrame.TextRange.Text) 'Contain the text as a string variable
                    
                    'Now that all text is saved as a string, check if the string itself completes all syntax requirements
                    'Try Table SX syntax first
                    If Left(currentText, 7) = "Table S" Then 'Check that text starts with "Table "
                        colonPosition = InStr(currentText, ":")
                        If colonPosition > 8 Then 'Ensure a ':' exists after the Table number
                            TablePrefix = Mid(currentText, 8, colonPosition - 8) 'extract the number
                            If IsNumeric(Trim(TablePrefix)) Then 'Verify the extracted value is numeric
                                TableNumber = CLng(Trim(TablePrefix)) 'save number as variable
                                remainingText = Mid(currentText, colonPosition) 'save remaining text as a String
                                
                                'Operations will now vary between the first slide and the subsequent slides.
                                If sld.SlideIndex = startSlideIndex And StartingTable = False Then 'For the first Table on the first slide only
                                    StartingTable = True
                                    prev_TableNumber = TableNumber
                                    'Print error if the slide does not contain a Table value
                                ElseIf StartingTable = False Then
                                        MsgBox "No Table caption found on the selected slide.", vbExclamation
                                        Exit Sub
                                'Still on the first slide, if there are multiple Tables on the slide that match the syntax
                                ElseIf sld.SlideIndex = startSlideIndex Then 'For the second Table on the first slide
                                    TableNumber = prev_TableNumber + 1
                                    shp.TextFrame.TextRange.Text = "Table S" & TableNumber & remainingText
                                    prev_TableNumber = TableNumber
                                
                                'For all remaining slides (because >= startSlideIndex was already established and we have now excluded =)
                                Else:
                                    'Increase the value
                                    TableNumber = prev_TableNumber + 1
                                    shp.TextFrame.TextRange.Text = "Table S" & TableNumber & remainingText
                                    prev_TableNumber = TableNumber
                                End If
                            End If
                        End If
                    
                    'Try Supplemental Table X syntax next
                    ElseIf Left(currentText, 19) = "Supplemental Table " Then 'Check that text starts with "Table "
                        colonPosition = InStr(currentText, ":")
                        If colonPosition > 21 Then 'Ensure a ':' exists after the Table number
                            TablePrefix = Mid(currentText, 20, colonPosition - 20) 'extract the number
                            If IsNumeric(Trim(TablePrefix)) Then 'Verify the extracted value is numeric
                                TableNumber = CLng(Trim(TablePrefix)) 'save number as variable
                                remainingText = Mid(currentText, colonPosition) 'save remaining text as a String
                                
                                'Operations will now vary between the first slide and the subsequent slides.
                                If sld.SlideIndex = startSlideIndex And IsEmpty(StartingTable) Then 'For the first Table on the first slide only
                                    StartingTable = TableNumber
                                    prev_TableNumber = StartingTable
                                    'Print error if the slide does not contain a Table value
                                    If IsEmpty(StartingTable) Then
                                            MsgBox "No Table caption found on the selected slide.", vbExclamation
                                            Exit Sub
                                    End If
                                'Still on the first slide, if there are multiple Tables on the slide that match the syntax
                                ElseIf sld.SlideIndex = startSlideIndex Then 'For the second Table on the first slide
                                    TableNumber = prev_TableNumber + 1
                                    shp.TextFrame.TextRange.Text = "Supplemental Table " & TableNumber & remainingText
                                    prev_TableNumber = TableNumber
                                
                                'For all remaining slides (because >= startSlideIndex was already established and we have now excluded =)
                                Else:
                                    'Increase the value
                                    TableNumber = prev_TableNumber + 1
                                    shp.TextFrame.TextRange.Text = "Supplemental Table " & TableNumber & remainingText
                                    prev_TableNumber = TableNumber
                                End If
                            End If
                        End If
                    
                    End If
                End If
            End If
        Next shp
    End If
Next sld

StartingTable = Empty

End Sub
