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

' exportFirstRow_Click
Sub exportFirstRow_Click()
    Call exportMain(1)
End Sub

' exportAllRows_Click
Sub exportAllRows_Click()
    Call exportMain(0) ' 0 means max
End Sub

' 0 means max; > 0 means the specified rows
Function exportMain(maxRowCount As Long)
    On Error GoTo Err_Handler

    Dim dataFilePath As String
    Dim templateDocPath As String
    Dim chosenDirBase As String
    Dim outputDirBase As String
    
    dataFilePath = showChooseExcelSourceDialog()
    If dataFilePath = "" Then
        Exit Function
    End If
    templateDocPath = showChooseDocTemplateDialog()
    If templateDocPath = "" Then
        Exit Function
    End If
    
    Dim isDirMade As Boolean
    isDirMade = False
    ' check and create output directory
    While Not isDirMade
        chosenDirBase = showChooseOutputBaseDir()
        If chosenDirBase = "" Then
            Exit Function
        End If
        
        outputDirBase = chosenDirBase & "\" _
            & "输出_" & Format(DateTime.Now, "yyyy-MM-dd_hh-mm-ss")   ' use "2018-04-12_15-34-56 etc
        isDirMade = makeDirectory(outputDirBase)
        If Not isDirMade Then
            MsgBox ("所选择的输出目录不存在，或者不能作为输出目录：" & vbCr & vbCr & "  " _
                & chosenDirBase & vbCr & vbCr _
                & "请选择另外的目录。")
        End If
    Wend
    
    ' output
    Call exportExcelToWordFiles(dataFilePath, templateDocPath, outputDirBase, maxRowCount)
    
    Application.ScreenUpdating = True
    MsgBox "完成。" & vbCr & vbCr & "请检查这个目录下的文件：" & vbCr & "  " & outputDirBase
    Exit Function

Err_Handler:
    Application.ScreenUpdating = True
    MsgBox "出现未知错误。"
    MsgBox "出现未知错误！" & vbCr & vbCr & "请检查这个目录下的文件：" & vbCr & "  " & outputDirBase

End Function

' 0 means max; > 0 means the specified rows
Function exportExcelToWordFiles(dataFilePath As String, _
    templateDocPath As String, _
    outputDirBase As String, _
    maxRowCount As Long)

On Error GoTo Err_Handler
    
    Dim wordApp As Word.Application
    Dim myDoc As Word.Document
    
    Dim dataBook As Workbook
    Dim dataSheet As Worksheet
    Set dataBook = Workbooks.Open(dataFilePath)
    Set dataSheet = dataBook.ActiveSheet
        
    Dim keepGoing As Boolean
    Dim rowIndex As Long
    Dim counter As Long
    Dim outputFilePath As String
    
    Set wordApp = New Word.Application
    wordApp.ScreenUpdating = False
    rowIndex = 2 ' from 3. #1 is the controlling tags (like $NAME), #2 is the captions
    counter = 0
    keepGoing = True
    Do While keepGoing
        ' prepare
        counter = counter + 1
        rowIndex = rowIndex + 1
        
        If Trim(dataSheet.Cells(rowIndex, 1)) <> "" Then ' check 1st column of each row
            ' prepare the new file
            Set myDoc = wordApp.Documents.Open(Filename:=templateDocPath)
            outputFilePath = outputDirBase & "\" & "Doc_" & CStr(counter + 1000) & ".docx"
            myDoc.SaveAs (outputFilePath) ' save as
            
            ' replacing
            Call innerHandleOneRow(dataSheet, rowIndex, myDoc)
            Call myDoc.Save
            Call myDoc.Close
            Set myDoc = Nothing
        Else
            keepGoing = False ' ends at empty line
        End If
        
        ' checking
        If maxRowCount > 0 And counter >= maxRowCount Then
            keepGoing = False
        End If
    Loop
    
    ' continue and clean up

Err_Handler:
    If Not wordApp Is Nothing Then
        wordApp.Quit
    End If
    If Not dataBook Is Nothing Then
        dataBook.Close
    End If
    If Not myDoc Is Nothing Then
        myDoc.Close
    End If
    
End Function

Function innerHandleOneRow(dataSheet As Worksheet, _
    rowIndex As Long, _
    myDoc As Word.Document)
    
    Dim col As Long
    Dim cell As Object
    Dim value As String
    
    col = 1
    Do While True
        Set cell = dataSheet.Cells(1, col) ' controlling row
        value = Trim(cell.value)
        
        If value = "" Then
            Exit Function
        End If
        
        If value Like "$*" Then
            Set cell = dataSheet.Cells(rowIndex, col) ' the data row
            
            Call replaceStringInDocument(myDoc, value, cell.value)
        End If
        
        ' next column
        col = col + 1
    Loop

    Exit Function

Err_Handler:
    
    
End Function
