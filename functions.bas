Public Function getWorkbookName() As String

    getWorkbookName = Application.ActiveWorkbook.FullName

End Function

Public Function getSQLString(inputText As String, Optional includeCommas As Boolean = True) As String

    Dim endingComma As String: endingComma = ""
    If includeCommas Then endingComma = ","
    
    getSQLString = "'" & Application.WorksheetFunction.Substitute(inputText, "'", "''") & "'" & endingComma
    
End Function

Public Function toJSON(ParamArray vals() As Variant) As String

    Dim i As Integer
    Dim workingArray As Variant
    Dim workingString As String: workingString = "{"
    
    'We need to have matching key / value pairs so the parameter needs to have an even number of args
    If (UBound(vals) + 1) Mod 2 <> 0 And UBound(vals) <> 0 Then Err.Raise vbObjectError + 1, "ExcelUtilities_Functions.getJSON", "This function requires an even number of arguments but an odd number was supplied."
    
    'if the inbound array is two-dimensional it's because another function passed it in as an array
    If UBound(vals) = 0 Then
        'Re-map the two-dimensional array to our one-dimensional
        ReDim workingArray(UBound(vals(0)))
        For i = LBound(vals(0)) To UBound(vals(0))
            workingArray(i) = vals(0)(i)
        Next i
    Else
        'Existing array is one-dimensional, pass it along unchanged.
        workingArray = vals
    End If
    
    'Build the output string
    For i = LBound(workingArray) To UBound(workingArray) Step 2
        workingString = workingString & Chr(34) & workingArray(i) & Chr(34) & ":" & Chr(34) & Replace(workingArray(i + 1), Chr(34), "\" & Chr(34)) & Chr(34) & ","
    Next i
    
    'Remove the last comma and add a closing brace
    workingString = Left(workingString, Len(workingString) - 1) & "}"
    
    toJSON = workingString

End Function

Public Function toJSONWithHeaders(selectedCells As Range) As String

    Dim i As Integer: i = 0
    Dim cell As Range
    Dim outboundArray() As Variant
    
    ReDim outboundArray((selectedCells.Count * 2) - 1)
    
    For Each cell In selectedCells
    
        outboundArray(i) = ActiveSheet.Cells(1, cell.Column).Value
        i = i + 1
        outboundArray(i) = cell.Value
        i = i + 1
    
    Next cell
    
    toJSONWithHeaders = toJSON(outboundArray)

End Function


Public Function getUUID() As String

    Dim i As Integer
    Dim binaryString As String: binaryString = getUUIDBinary()
    Dim workingString As String: workingString = ""
    
    'This mid function uses 1 as its base, so starting at 1 instead of zero to reuse variable
    For i = 1 To 128 Step 4
    
        workingString = workingString & Application.WorksheetFunction.Bin2Hex(Mid(binaryString, i, 4))
    
    Next i
    
    getUUID = LCase(Format(workingString, "&&&&&&&&-&&&&-&&&&-&&&&-&&&&&&&&&&&&"))

End Function

Private Function getUUIDBinary() As String

    Dim i As Integer
    Dim binaryString As String: binaryString = ""
    Dim binaryDigit As Integer
    
    Randomize
    
    For i = 0 To 127
    
        Select Case i
        
            Case 48, 50, 51, 65
                binaryDigit = 0
            
            Case 49, 64
                binaryDigit = 1
                
            Case Else
                binaryDigit = CInt(1 * Rnd)
        
        End Select
        
        binaryString = binaryString & binaryDigit
    
    Next i
    
    getUUIDBinary = binaryString

End Function
