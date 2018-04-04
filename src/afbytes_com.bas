Attribute VB_Name = "afbytes_com"
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

' #### Import CSV, and convert the format, and store it in a temporary Excel file
Sub ImportCSV_Button_Click()
Attribute ImportCSV_Button_Click.VB_ProcData.VB_Invoke_Func = "g\n14"
    On Error GoTo Err_Handler
    
    Dim csvFilePath As String
    csvFilePath = showChooseCSVDialog()
    If csvFilePath = "" Then
        Exit Sub
    End If
    
    ' ---- import
    Application.ScreenUpdating = False
    
    Dim tempWb As Workbook
    Dim tempSrcWs As Worksheet
    Dim tempOutWs As Worksheet
    
    Set tempWb = Workbooks.Add
    Set tempOutWs = tempWb.Sheets(1) ' output to the 1st sheet
    Set tempSrcWs = tempWb.Sheets(2) ' store to the second sheet
        
    With tempSrcWs.QueryTables.Add(Connection:="TEXT;" & csvFilePath, Destination:=tempSrcWs.Range("A1"))
         .TextFilePlatform = 65001 ' consider it as UTF-8
         .TextFileParseType = xlDelimited
         .TextFileCommaDelimiter = True
         .Refresh
    End With
    
    tempOutWs.UsedRange.ClearContents ' clear content of a sheet
    Call innerHandleData(tempOutWs, tempSrcWs)
    
    Application.ScreenUpdating = True
    
    ' show tip
    Dim tipText As String
    tipText = "转换成功。" & vbCr & vbCr _
        & "你可以：" & vbCr _
        & "  (A) 保存这个新建的Excel文件，或者只是" & vbCr _
        & "  (B) 拷贝其中内容（无须保存新建的文件）。"
    Call MsgBox(tipText, vbOKOnly, "AFBytes Studio")
    
    Exit Sub

Err_Handler:
    Application.ScreenUpdating = True
    MsgBox "出现未知错误。"
End Sub

' return the full file path, or empty string when no selection
Function showChooseCSVDialog()
    Dim strFileToOpen ' As String
    
    Dim title As String: title = "请选择待打开的CSV文件："
    Dim filter As String: filter = _
        "CSV 文件 (*.csv), *.csv"
    Dim allowMultiple As Variant: allowMultiple = False ' use False when using "Option Explicit"
    
    strFileToOpen = Application.GetOpenFilename( _
        title:=title, FileFilter:=filter, MultiSelect:=allowMultiple)
        
    If strFileToOpen = "False" Then ' "False" is returned when no file chosen
        showChooseCSVDialog = ""
        Exit Function
    End If
    
    ' your action
    showChooseCSVDialog = strFileToOpen
End Function

Rem ---------------------------------------------------------------------------------------------------

' Iterate all columns, and merge the answers
Function innerHandleData(outputSheet As Worksheet, sourceSheet As Worksheet)
    ' outputSheet.UsedRange.Columns.ColumnWidth = 100 ' make columns large enough
    outputSheet.Cells.EntireColumn.ColumnWidth = 100  ' make columns large enough
    
    ' handle data
    Dim outputBase As Long: outputBase = 10
    Dim colIdx As Long
    Dim rowIdx As Long
    Dim prevColCaption As String: prevColCaption = ""
    colIdx = 1
    rowIdx = 1
    
    Dim cell As Object
    Dim colCaption As String
    Dim realCaption As String
    
    Dim outputColumnIndex As Long
    Dim lastSourceRowId As Long
    lastSourceRowId = getEndingRow(sourceSheet, 1, 1)
    
    outputColumnIndex = 0 ' begin from invalid value 0
    Do While True
        Set cell = sourceSheet.Cells(rowIdx, colIdx)
        
        colCaption = cell.Value
        If Trim(colCaption) = "" Or colIdx >= 5000 Then
            Exit Do ' end when meet empty cell
        End If
        
        ' get real caption
        realCaption = getCaptionOfMultiChoicesQuestion(colCaption)
           
        ' copy the value
        'outputSheet.Cells(outputBase + colIdx, 1).Value = cell.Value
        
        ' determine Mutiple Choices questions
        Dim opType As String
        If realCaption <> "" And realCaption = prevColCaption Then
            opType = "append"
        Else
            opType = "copy"
        End If
        ' determine Customized-Answer field
        If endsWith(colCaption, "[选项填空]") Then ' this condition override the previous one
            opType = "concat"
        End If
        
        If opType = "copy" Then ' next output column
            outputColumnIndex = outputColumnIndex + 1
        End If
        
        ' settle the column
        Rem Debug.Print CStr(colIdx) & " -> " & CStr(outputColumnIndex)
        Call outputDataColumn(outputSheet, outputColumnIndex, realCaption, opType, _
                              sourceSheet, colIdx, lastSourceRowId)
        
        ' next column
        colIdx = colIdx + 1
        prevColCaption = realCaption
    Loop
    
    ' auto fit the output
    outputSheet.UsedRange.Columns.AutoFit

End Function


' return: the row id after the last non-empy row
Function getEndingRow(ByRef Worksheet As Worksheet, col As Long, fromRow As Long) As Long
    Dim cell As Object
    Dim row As Long
    
    row = fromRow
    Do While True
        Set cell = Worksheet.Cells(row, col)
        If Trim(cell.Value) = "" Then
            Exit Do
        End If
        
        row = row + 1
    Loop
    
    getEndingRow = row

End Function

' return the original or filtered caption according to Tencent WJ's format
Function getCaptionOfMultiChoicesQuestion(caption As String) As String
    If caption = "" Then
        getCaptionOfMultiChoicesQuestion = ""
        Exit Function
    End If
    
    Dim strPattern As String: strPattern = "^([0-9]+\..*):.+"   ' -- 1.Which Fruit:Apple
    Dim regEx As New RegExp

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern ' the RegEx string
    End With

    If regEx.Test(caption) Then
        Dim matches
        Set matches = regEx.Execute(caption)
        getCaptionOfMultiChoicesQuestion = matches(0).SubMatches(0)
    Else
        getCaptionOfMultiChoicesQuestion = caption ' return the orignal one
    End If

End Function

' operationType: "copy", "append", "concat"
Sub outputDataColumn(outputSheet As Worksheet, _
                     columnIndex As Long, _
                     columnCaption As String, _
                     operationType As String, _
                     ByRef sourceSheet As Object, _
                     sourceColumnIndex As Long, _
                     sourceEndingRowIndex As Long)
    Dim n As Long
    
    Dim srcCell As Object
    Dim dstCell As Object
    
    Dim rowIdx As Long
    
    ' question captions
    rowIdx = 1
    Set dstCell = outputSheet.Cells(rowIdx, columnIndex)
    If operationType = "copy" Then
        dstCell.Value = columnCaption
    End If

    ' answers
    Dim srcText As String
    For rowIdx = 2 To sourceEndingRowIndex - 1
        Set srcCell = sourceSheet.Cells(rowIdx, sourceColumnIndex)
        Set dstCell = outputSheet.Cells(rowIdx, columnIndex)
        
        srcText = Trim(srcCell.Value)
        If srcText <> "" Then
            If operationType = "copy" Then
                dstCell.Value = srcText
            ElseIf operationType = "append" Then
                If dstCell.Value <> "" Then
                    dstCell.Value = dstCell.Value & Chr(13) & Chr(10) & srcText
                Else
                    dstCell.Value = srcText
                End If
            ElseIf operationType = "concat" Then
                dstCell.Value = dstCell.Value & srcText
            Else
                Exit Sub ' unsupported type, just exist
            End If
        End If
    Next rowIdx
End Sub

