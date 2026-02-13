' ============================
' Convert DOCX to PDF using Word
' ============================
Dim wordApp
Dim doc

inDocPath = WScript.Arguments(0)
inPdfPath = WScript.Arguments(1)

On Error Resume Next

Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

Set doc = wordApp.Documents.Open(inDocPath)

' 17 = wdExportFormatPDF
doc.ExportAsFixedFormat inPdfPath, 17

doc.Close False
wordApp.Quit

If Err.Number = 0 Then
    outStatus = "SUCCESS"
Else
    outStatus = "ERROR: " & Err.Description
End If

Set doc = Nothing
Set wordApp = Nothing
