Attribute VB_Name = "Automatic_TableofContents"
Option Explicit

Sub Auto_Table_Contents()

    Dim StartCell As Range 'for input box to select range
    Dim Sh As Worksheet
    Dim ShName As String
    Dim MsgConfirm As VBA.VbMsgBoxResult ' for message box to confirm
    Dim EndCell As Range
    
    
    On Error Resume Next
    
    Set StartCell = Application.InputBox("Where do you want to insert the Table of Contents?" _
                & vbNewLine & "Please select a cell:", "Insert Table of Contents", , , , , , 8)
    
    If Err.Number = 424 Then Exit Sub
    On Error GoTo Handle
    Set StartCell = StartCell.Cells(1, 1)
    Set EndCell = StartCell.Offset(Worksheets.Count - 2, 1)
    
    'get confirmation
    MsgConfirm = VBA.MsgBox("The value in cells: " & vbNewLine & StartCell.Address & " to " & EndCell.Address & " could be overwritten." & vbNewLine & "Would you like to continue?", vbOKCancel + vbDefaultButton2, "Confirmation Required!")
    If MsgConfirm = vbCancel Then Exit Sub
    
    'Loop through each worksheet and extract sheet name
    For Each Sh In Worksheets
        ShName = Sh.Name
        'Exclude the activesheet from the table of content
        If ActiveSheet.Name <> ShName Then
            'Only add visible sheets
            If Sh.Visible = xlSheetVisible Then
                'Add a hyperlink to the sheet name value
                ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:= _
                "'" & ShName & "'" & "!A1", TextToDisplay:=ShName '"'" & ShName & "'" & "!A1": sub address ShName was put in single quotations so that the name of the Tabs are copied exactly to how it is
                StartCell.Offset(0, 1).Value = Sh.Range("A1").Value
                Set StartCell = StartCell.Offset(1, 0)
            End If 'sheet is visible
        End If 'sheet is not active sheet
    Next Sh
Exit Sub
Handle:
MsgBox "Unfortunately an error has occurred"

End Sub
