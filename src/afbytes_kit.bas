' Copyright (c) 2018 AFBytes Studio.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

Option Explicit

' https://www.afbytes.com
'MIT License
' -------------------------------------------------------------------------------

' return: the row id after the last non-empy row
Function getEndingRow(ByRef Worksheet As Worksheet, col As Long, fromRow As Long) As Long
    Dim cell As Object
    Dim row As Long
    
    row = fromRow
    Do While True
        Set cell = Worksheet.Cells(row, col)
        If Trim(cell.value) = "" Then
            Exit Do
        End If
        
        row = row + 1
    Loop
    
    getEndingRow = row

End Function


Function replaceStringInDocument(myDoc As Word.Document, _
                               fromStr As String, toStr As String)
    Dim rngStory
    For Each rngStory In myDoc.StoryRanges
        With rngStory.Find
          .text = fromStr
          .Replacement.text = toStr
          .Wrap = wdFindContinue
          If .Execute(Replace:=wdReplaceAll) Then
            ' whatChanged = sFileName & "|" & strFind & "|" & strReplace & "|" & Now()
            ' Print #FileNum, whatChanged
            ' Debug.Print fromStr & " -> " & toStr & " | " & Now()
          End If
        End With
    Next rngStory

End Function

' return True when the path points to an existing directory
'   Note: "C:", "C:\" result in False.
Function isDirectoryAndExist(path As String)
    On Error GoTo ErrorHandler
    
    Dim resultValue As Boolean
    resultValue = False
    
    ' GetAttr will raise exception when target doesn't exist
    If GetAttr(path) = vbDirectory Then
        resultValue = True
    End If
    
    isDirectoryAndExist = resultValue
    Exit Function
    
ErrorHandler:
    isDirectoryAndExist = False

End Function

' return True when created successfully, otherwise False
Function makeDirectory(path As String)
    On Error GoTo finalStep
    
    Dim result As Boolean
    result = False
    
    MkDir (path)
    result = True
    
finalStep:
    makeDirectory = result

End Function
