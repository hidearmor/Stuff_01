Attribute VB_Name = "Donez"
Function CreateNewTable() As Table
    ' Create and return a new table
    Set CreateNewTable = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=4)
    With CreateNewTable
        .Rows.Height = CentimetersToPoints(6.8)
        .Columns.Width = CentimetersToPoints(4.4)
        .Borders.Enable = True  ' Enable table borders
        .Rows.Alignment = wdAlignRowCenter
    End With
End Function

Sub CreateTableFromText()

        ' Declare variables
    Dim originalDoc As Document
    Dim newDoc As Document

    ' Set the original document (currently active)
    Set originalDoc = ActiveDocument
    
        ' Get user input
    Dim userInput As String
    userInput = ActiveDocument.Range.text

    ' Create a new document
    Set newDoc = Documents.Add
    
        ' Set the page setup of the new document to A4
    With newDoc.PageSetup
        .PaperSize = wdPaperA4
    End With

    ' Make the new document the active document
    newDoc.Activate
    
    '        Set document margins
        With ActiveDocument.PageSetup
            .LeftMargin = CentimetersToPoints(1.27)
            .RightMargin = CentimetersToPoints(1.27)
            .TopMargin = CentimetersToPoints(1)
            .BottomMargin = CentimetersToPoints(1)
        End With

    ActiveDocument.Content.Delete

    ' Check if user canceled
    If Len(userInput) = 0 Then
        Exit Sub
    End If

    ' Split text by *
    Dim words() As String
    words = Split(userInput, "*")

    ' Set up variables
    Dim currentRow As Integer
    Dim currentCol As Integer
    Dim tbl As Table
    currentRow = 1
    currentCol = 1
    Dim wordIndex As Integer
    Dim wordIndex2 As Integer
    wordIndex = 0
    wordIndex2 = 1
    Dim last As Boolean
    last = False
    
    Do While wordIndex2 <= (UBound(words))

'     find length of the loop
        Dim length As Integer
        If UBound(words) - wordIndex > 31 Then
            length = 15
        Else
            last = True
            length = ((UBound(words) - wordIndex) \ 2)
        End If
        
        Set tbl = CreateNewTable
        currentCol = 1
        
        For i = 0 To length
            If wordIndex <= UBound(words) Then
                tbl.Cell(currentRow, currentCol).VerticalAlignment = wdCellAlignVerticalCenter
'                tbl.Cell(currentRow, currentCol).Range.text = wordIndex & " " & words(wordIndex)
                tbl.Cell(currentRow, currentCol).Range.text = words(wordIndex)

                ' Set text properties
                With tbl.Cell(currentRow, currentCol).Range
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter  ' Center horizontally
                    .ParagraphFormat.SpaceBefore = 0
                    .ParagraphFormat.SpaceAfter = 0
                    .Cells.VerticalAlignment = wdCellAlignVerticalCenter  ' Center vertically
                End With

                ' Move to the next cell
                currentCol = currentCol + 1
                If currentCol > 4 Then
                    currentCol = 1
                    currentRow = currentRow + 1
                End If

                ' Move to the next word
                wordIndex = wordIndex + 2
            End If
        Next i
        
        If last = True Then
            Selection.EndKey Unit:=wdStory
'            Selection.InsertBreak Type:=wdPageBreak
            Set tbl = CreateNewTable
'            currentRow = 1
            currentRow = currentRow + ((length + 1) Mod 4)
        Else
            Selection.EndKey Unit:=wdStory
            Set tbl = CreateNewTable
        End If
        
        
        currentCol = 4
        
        For i = 0 To length
            If wordIndex2 <= UBound(words) Then
                tbl.Cell(currentRow, currentCol).VerticalAlignment = wdCellAlignVerticalCenter
'                tbl.Cell(currentRow, currentCol).Range.text = wordIndex2 & " " & words(wordIndex2)
                tbl.Cell(currentRow, currentCol).Range.text = words(wordIndex2)

                ' Set text properties
                With tbl.Cell(currentRow, currentCol).Range
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter  ' Center horizontally
                    .ParagraphFormat.SpaceBefore = 0
                    .ParagraphFormat.SpaceAfter = 0
                    .Cells.VerticalAlignment = wdCellAlignVerticalCenter  ' Center vertically
                End With

                ' Move to the next cell
                currentCol = currentCol - 1
                If currentCol < 1 Then
                    currentCol = 4
                    currentRow = currentRow + 1
                End If
                

                ' Move to the next word
                wordIndex2 = wordIndex2 + 2
            End If
        Next i
        
        If wordIndex <= UBound(words) Then
            Selection.EndKey Unit:=wdStory
        End If
        If wordIndex2 >= UBound(words) Then
            Exit Do
        End If
    Loop
    
    Dim desiredFontSize As Single
    desiredFontSize = 22 ' Change this to your desired size
    For Each para In newDoc.Paragraphs
        para.Range.Font.Size = desiredFontSize
    Next para
End Sub


