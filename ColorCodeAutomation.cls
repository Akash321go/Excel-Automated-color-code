VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ColorCodeRowsInAnyFileAndSave()
    Dim fd As FileDialog
    Dim FileChosen As Integer
    Dim FileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim WorkOrderRange As Range
    Dim CurrentDataRange As Range
    Dim cell As Range
    Dim WorkOrderDict As Object
    Dim ColorIndex As Integer
    Dim LightColors As Variant
    Dim LastColumn As Integer
    
    ' Define light colors using RGB values for better readability
    LightColors = Array(RGB(255, 255, 204), RGB(218, 238, 243), RGB(242, 220, 219), _
                        RGB(221, 235, 247), RGB(237, 227, 210), RGB(204, 236, 255), _
                        RGB(252, 229, 205), RGB(247, 202, 201))
    
    ' Create a FileDialog object as a File Picker dialog box
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Use the Show method to display the File Picker dialog box and return the user's action
    FileChosen = fd.Show
    
    If FileChosen = -1 Then
        ' Get the path to the file chosen by the user
        FileName = fd.SelectedItems(1)
        
        ' Open the chosen workbook
        Set wb = Workbooks.Open(FileName)
        
        ' Set the worksheet you want to color code, for now, we'll use the first sheet
        Set ws = wb.Sheets(1) 'Change this if you need to use a specific sheet
        
        ' Assuming the Work Order IDs are in column A, and data starts from row 2
        ' Modify this range according to the structure of your workbooks
        Set WorkOrderRange = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
        LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Initialize the dictionary and color index
        Set WorkOrderDict = CreateObject("Scripting.Dictionary")
        ColorIndex = 0
        
        ' Apply color-coding to the data range
        For Each cell In WorkOrderRange
            If cell.Value <> "" Then
                If Not WorkOrderDict.exists(cell.Value) Then
                    WorkOrderDict.Add cell.Value, LightColors(ColorIndex)
                    ColorIndex = (ColorIndex + 1) Mod (UBound(LightColors) + 1)
                End If
                Set CurrentDataRange = ws.Range(cell, ws.Cells(cell.Row, LastColumn))
                CurrentDataRange.Interior.Color = WorkOrderDict(cell.Value)
                With CurrentDataRange.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
            End If
        Next cell
        
        ' Autofit the columns for the entire data range
        ws.Columns("A:H").AutoFit ' Change the range accordingly if your data size is different
        
        ' Save the workbook
        wb.Save
        
        ' Inform the user
        MsgBox "Color coding applied and file saved successfully to " & wb.Name, vbInformation
        
        ' Close the workbook
        wb.Close
        
    Else
        ' The user pressed "Cancel"
        MsgBox "No file was selected.", vbExclamation
    End If
    
    ' Clean up
    Set fd = Nothing
    Set WorkOrderDict = Nothing
End Sub


