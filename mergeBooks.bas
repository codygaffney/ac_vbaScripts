Attribute VB_Name = "Module2"
Sub mergeFiles()
    'Merges all files in a folder to a main file.
    
    'Define variables:
    Dim numberOfFilesChosen, i As Integer
    Dim tempFileDialog As FileDialog
    Dim mainWorkbook, sourceWorkbook As Workbook
    Dim tempWorkSheet As Worksheet
    
    Set mainWorkbook = Application.ActiveWorkbook
    Set tempFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    'Allow the user to select multiple workbooks
    tempFileDialog.AllowMultiSelect = True
    
    numberOfFilesChosen = tempFileDialog.Show
    
    'Loop through all selected workbooks
    For i = 1 To tempFileDialog.SelectedItems.Count
        
        'Open each workbook
        Workbooks.Open tempFileDialog.SelectedItems(i)
        
        Set sourceWorkbook = ActiveWorkbook
        
        'Copy each worksheet to the end of the main workbook
        For Each tempWorkSheet In sourceWorkbook.Worksheets
            tempWorkSheet.Copy after:=mainWorkbook.Sheets(mainWorkbook.Worksheets.Count)
        Next tempWorkSheet
        
        'Close the source workbook
        sourceWorkbook.Close
    Next i
    
End Sub
