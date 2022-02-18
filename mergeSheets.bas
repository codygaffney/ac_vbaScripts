Attribute VB_Name = "Module1"
Sub Append_Data_From_Different_Sheets_Into_Single_Sheet_By_Column()
'Procedure to Consolidate all sheets in a workbook

On Error GoTo IfError

'1. Variables declaration
Dim Sht As Worksheet, DstSht As Worksheet
Dim LstRow As Long, LstCol As Long, DstCol As Long
Dim i As Integer, EnRange As String
Dim SrcRng As Range

'2. Disable Screen Updating - stop screen flickering
'   And Disable Events to avoid inturupted dialogs / popups
With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

'3. Delete the Append_Data WorkSheet if it exists
Application.DisplayAlerts = False
On Error Resume Next
ActiveWorkbook.Sheets("Append_Data").Delete
Application.DisplayAlerts = True

'4. Add a new WorkSheet and name as 'Append_Data'
With ActiveWorkbook
    Set DstSht = .Sheets.Add(After:=.Sheets(.Sheets.Count))
    DstSht.Name = "Append_Data"
End With

'5. Loop through each WorkSheet in the workbook and copy the data to the 'Append_Data' WorkSheet
For Each Sht In ActiveWorkbook.Worksheets
    If Sht.Name <> DstSht.Name Then
       '5.1: Find the last row on the 'Append_Data' sheet
       Rem DstCol = fn_LastColumn(DstSht)
       DstRow = fn_LastRow(DstSht)
           
       If DstCol = 1 Then DstCol = 0
               
       '5.2: Find Input data range
       LstRow = fn_LastRow(Sht)
       LstCol = fn_LastColumn(Sht)
       EnRange = Sht.Cells(LstRow, LstCol).Address
       Set SrcRng = Sht.Range("A2:" & EnRange)
       
       '5.3: Check whether there are enough columns in the 'Append_Data' Worksheet
        If DstCol + SrcRng.Columns.Count > DstSht.Columns.Count Then
            MsgBox "There are not enough columns to place the data in the Append_Data worksheet."
            GoTo IfError
        End If
                
      '5.4: Copy data to the 'Append_Data' WorkSheet
        SrcRng.Copy Destination:=DstSht.Cells(DstRow + 1, 1)
    End If
Next
DstSht.Range("A1") = "Customer"
DstSht.Range("B1") = "Customer ID"
DstSht.Range("C1") = "Job Name"
DstSht.Range("D1") = "Material"
DstSht.Range("E1") = "Notes"
DstSht.Range("F1") = "Qty"
DstSht.Range("G1") = "Area"
DstSht.Range("H1") = "Rate"

IfError:
'6. Enable Screen Updating and Events
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub


'In this example we are finding the last Row of specified Sheet
'In this example we are finding the last Row of specified Sheet
Function fn_LastRow(ByVal Sht As Worksheet)

    Dim lastRow As Long
    lastRow = Sht.Cells.SpecialCells(xlLastCell).Row
    lRow = Sht.Cells.SpecialCells(xlLastCell).Row
    Do While Application.CountA(Sht.Rows(lRow)) = 0 And lRow <> 1
        lRow = lRow - 1
    Loop
    fn_LastRow = lRow

End Function

'In this example we are finding the last column of specified Sheet
Function fn_LastColumn(ByVal Sht As Worksheet)

    Dim lastCol As Long
    lastCol = Sht.Cells.SpecialCells(xlLastCell).Column
    lCol = Sht.Cells.SpecialCells(xlLastCell).Column
    Do While Application.CountA(Sht.Columns(lCol)) = 0 And lCol <> 1
        lCol = lCol - 1
    Loop
    fn_LastColumn = lCol

End Function
