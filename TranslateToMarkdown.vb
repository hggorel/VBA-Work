Sub TranslateMarkdown()

    Dim cell As Range
    Dim selectedRange As Range
    
    Set selectedRange = Application.Selection
    
    Dim rowCounter As Integer
    Dim colCounter As Integer

    Dim totCol As Integer
    Dim totRow As Integer
    
    totCol = selectedRange.Columns.Count

    '/// Setting the lines exactly
    Dim thisLine As String
    Dim Color As String
    Dim checkColor As String
    Dim colorCompare As Integer
    Dim linkCount As Integer
    Dim link As String
    Dim includeLink As Boolean
    Dim isBold As Boolean
    
    Dim col As Integer
    Dim currRow As Integer
    
    Dim numColHidden As Integer
    
    rowCounter = 0
    numColHidden = 0
    
    For Each row In selectedRange.Rows
        
        colCounter = 0
        
        thisLine = "|"
        
        For Each cell In row.Cells
            col = cell.Column
            currRow = cell.row
            
            '///Debug.Print (Columns(col).Hidden)
            If Columns(col).Hidden = False And Rows(currRow).Hidden = False Then
                
                includeLink = False
                isBold = False
        
                '/// Check to see what the color is and if yellow - ignore
            
                Color = cell.Interior.Color
                checkColor = Hex(Color)
            
                colorCompare = StrComp("FFFF", checkColor)
            
                '/// Check for hyper link
            
                linkCount = cell.Hyperlinks.Count
            
                isBold = cell.Font.Bold
                        
                If linkCount > 0 Then
                    Dim firstChar As String
                    Dim charCompare As String
                    firstChar = Left$(cell.Value, 1)
                    charCompare = StrComp(firstChar, "!")
                
                    If charCompare = 0 Then
                        includeLink = False
                    Else
                        includeLink = True
                        link = cell.Hyperlinks(1).Address
                    End If
                End If
            
                If (colorCompare = 0) Then
                    thisLine = thisLine & " "
                Else
            

            
                    thisLine = thisLine & " "
                
                    If includeLink Then
                        If isBold Then
                            thisLine = thisLine & "[**"
                            thisLine = thisLine & cell.Value
                            thisLine = thisLine & "**]("
                        Else
                            thisLine = thisLine & "["
                            thisLine = thisLine & cell.Value
                            thisLine = thisLine & "]("
                        End If
                    
                        thisLine = thisLine & link
                        thisLine = thisLine & ")"
                    Else
                        If isBold Then
                            thisLine = thisLine & "**"
                        End If
                        thisLine = thisLine & cell.Value
                        If isBold Then
                            thisLine = thisLine & "**"
                        End If
                    
                    End If
                
                    thisLine = thisLine & "|"
            End If
            colCounter = colCounter + 1
            Else
                numColHidden = numColHidden + 1
                
            End If
            Next cell
            If Rows(currRow).Hidden = False Then
                Debug.Print thisLine
            End If
        
            '/// Need to account for the lines to make it a table right after Row 0 is counted
        
            If (rowCounter = 0) Then
                thisLine = "|"
                colCounter = 0
            
                For i = 0 To (totCol - 1 - numColHidden)
                    thisLine = thisLine
                
                    thisLine = thisLine & "----|"
                    colCounter = colCounter + 1
                Next i
            
                Debug.Print thisLine
            End If
        
            rowCounter = rowCounter + 1
        
        Next row
        
    Debug.Print
    Debug.Print
    
    
End Sub
