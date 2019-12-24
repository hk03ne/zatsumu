Option Explicit

dim fs
dim excel
dim arg
dim wb
dim ws

Set fs = CreateObject("Scripting.FileSystemObject")
Set excel = CreateObject("Excel.Application")

For Each arg In WScript.Arguments
    If Not fs.FileExists(arg) Then 
        MsgBox("File not found: " & arg)
        WScript.Quit
    End If

    Set wb = excel.Workbooks.Open(arg)

    For Each ws In wb.Worksheets
        ws.Activate
        ws.Cells(1,1).Select
        excel.ActiveWindow.View = 2 ' xlPageBreakPreview
        excel.ActiveWindow.Zoom = 100
    Next

    wb.Worksheets(1).Activate
    wb.Save
    wb.Close
Next

