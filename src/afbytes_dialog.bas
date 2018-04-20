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

' return the full file path, or empty string when no selection
Function showChooseExcelSourceDialog()
    Dim strFileToOpen ' As String
    
    Dim title As String: title = "请选择Excel数据文件："
    Dim filter As String: filter = _
        "Excel 文件 (*.xls*), *.csv*"
    Dim allowMultiple As Variant: allowMultiple = False ' use False when using "Option Explicit"
    
    strFileToOpen = Application.GetOpenFilename( _
        title:=title, FileFilter:=filter, MultiSelect:=allowMultiple)
        
    If strFileToOpen = "False" Then ' "False" is returned when no file chosen
        showChooseExcelSourceDialog = ""
        Exit Function
    End If
    
    ' your action
    showChooseExcelSourceDialog = strFileToOpen
End Function

' return the full file path, or empty string when no selection
Function showChooseDocTemplateDialog()
    Dim strFileToOpen ' As String
    
    Dim title As String: title = "请选择Word模板文件："
    Dim filter As String: filter = _
        "Word 文件 (*.doc*), *.doc*"
    Dim allowMultiple As Variant: allowMultiple = False ' use False when using "Option Explicit"
    
    strFileToOpen = Application.GetOpenFilename( _
        title:=title, FileFilter:=filter, MultiSelect:=allowMultiple)
        
    If strFileToOpen = "False" Then ' "False" is returned when no file chosen
        showChooseDocTemplateDialog = ""
        Exit Function
    End If
    
    ' your action
    showChooseDocTemplateDialog = strFileToOpen
End Function

' return the full file path, or empty string when no selection
Function showChooseOutputBaseDir()
    Dim title As String: title = "请选择保存生成文件的目录（会自动创建子目录）："
    Dim allowMultiple As Variant: allowMultiple = False ' use False when using "Option Explicit"
    
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    sItem = ""
    With fldr
        .title = title
        .AllowMultiSelect = allowMultiple
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:
    showChooseOutputBaseDir = sItem
    Set fldr = Nothing

End Function
