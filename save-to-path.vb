Sub SaveAsPDF()
    On Error GoTo ErrHandler

    Dim InvoiceDate As String
    Dim Path As String
    Dim fileName As String
    Dim userFileName As String
    Dim serialNumber As Integer
    Dim fullFileName As String

    ' Define the file path
    Path = "C:\Users\user\OneDrive\Documents\OA Tutor\Docs\Finance\Invoice\All\"

    ' Prompt user to enter a custom file name
    ' userFileName = InputBox("Enter the file name (without extension):", "Save As PDF")

    ' If user cancels or enters an empty value, exit the subroutine
    ' If userFileName = "" Then
    '     MsgBox "No file name entered. Operation canceled."
    '     Exit Sub
    ' End If

    InvoiceDate = Format(Now, "MMDDYYYY") ' Use current date in MMDDYYYY format for InvoiceNo

    ' Set the starting serial number to 37
    serialNumber = 37

    ' Check if a file with the name already exists, and increment the serial number if it does
    Do While Dir(Path & InvoiceDate & " - Invoice " & Format(serialNumber, "000") & ".pdf") <> ""
        serialNumber = serialNumber + 1
    Loop

    ' Create the full file name with the serial number
    fullFileName = InvoiceDate & " - Invoice " & Format(serialNumber, "000") & ".pdf"

    ' Export the active sheet as a PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=Path & fullFileName, IgnorePrintAreas:=False

Exit Sub

ErrHandler:
    MsgBox "Error saving PDF: " & Err.Description
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
