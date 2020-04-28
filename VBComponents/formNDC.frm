VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNDC 
   Caption         =   "National Drug Code Directory"
   ClientHeight    =   3012
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3864
   OleObjectBlob   =   "formNDC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formNDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const frameFinished As String = "NDC finished products search"
Const frameUnfinished As String = "Unfinished Products"

Private Sub UserForm_Initialize()

    ' Set up userform to default
    optFinished.Value = True
    optFinished_Click
    frameSearch.Caption = frameFinished
    
    ' Set control tip text
    optFinished.ControlTipText = "FDA reviewed and approved products."
    optUnfinished.ControlTipText = "Unapproved products."
    comboType.ControlTipText = "Select a search type."
    textSearch.ControlTipText = "Select a search type."
    comboSheet.ControlTipText = "Select an output sheet."
    btnSearch.ControlTipText = "Click to search."
    btnClear.ControlTipText = "Click to clear the search text."
    
    ' Set up sheets combo box
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        comboSheet.AddItem ws.Name
    Next ws
    comboSheet.AddItem "(Create new sheet)"
    comboSheet.Value = "Select Sheet"
    
End Sub

Private Sub comboType_Change()
    Select Case comboType.Value
        Case "Brand Name"
            textSearch.ControlTipText = "Please type the full name of the drug. The API does not support partial searches."
        Case "Application Number"
            textSearch.ControlTipText = "Please type the application number of the drug."
        Case "Generic Name"
            textSearch.ControlTipText = "Please type the full name of the drug. The API does not support partial searches."
        Case "NDC"
            textSearch.ControlTipText = "Please type the manufacturer code and product code separated by a hyphen."
        Case "Labeler"
            textSearch.ControlTipText = "Please type the full name of the labeler. The API does not support partial searches."
    End Select
End Sub

Private Sub optFinished_Click()
    If optFinished.Value = True Then
        frameSearch.Caption = frameFinished
        With comboType
            .Clear
            .AddItem "Brand Name"
            .AddItem "Application Number"
            .AddItem "Generic Name"
            .AddItem "NDC"
            .AddItem "Labeler"
            .Value = "Select Type"
        End With
    End If
End Sub

Private Sub optUnfinished_Click()
    If optUnfinished.Value = True Then
        frameSearch.Caption = frameUnfinished
        With comboType
            .Clear
            .AddItem "Generic Name"
            .AddItem "NDC"
            .AddItem "Labeler"
            .Value = "Select Type"
        End With
    End If
End Sub

Private Sub btnSearch_Click()
    
    ' Validate inputs
    If comboType.ListIndex = -1 Then
        MsgBox "Select Type", vbInformation
        Exit Sub
    ElseIf Len(textSearch.Text) = 0 Then
        MsgBox "Add search text", vbInformation
        Exit Sub
    ElseIf comboSheet.ListIndex = -1 Then
        MsgBox "Select Sheet", vbInformation
        Exit Sub
    ElseIf InStr(1, textSearch.Text, "-") < 1 And _
        comboType.Value = "NDC" Then
        MsgBox "Please include a hyphen between manufacturer code and product code.", vbInformation
        Exit Sub
    End If
    
    ' Add new sheet if option is selected
    If comboSheet.ListIndex = comboSheet.ListCount - 1 Then
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    End If
    
    ' Run API
    If optFinished.Value = True Then
        Me.Hide
        Call API_NDC_PullData_Dict(textSearch.Text, comboType.ListIndex, True, _
            ActiveWorkbook.Worksheets(comboSheet.ListIndex + 1))
    Else
        Me.Hide
        Call API_NDC_PullData_Dict(textSearch.Text, comboType.ListIndex + 2, False, _
            ActiveWorkbook.Worksheets(comboSheet.ListIndex + 1))
    End If
    
    Unload Me
    
End Sub

Private Sub btnClear_Click()
    textSearch.Text = ""
End Sub
