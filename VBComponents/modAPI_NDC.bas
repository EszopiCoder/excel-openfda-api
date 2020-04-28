Attribute VB_Name = "modAPI_NDC"
Option Explicit
' API NDC directory constant
Private Const API_NDC_Directory = "https://api.fda.gov/drug/ndc.json?search="
' API search fields
Private Const API_Brand = "brand_name:"
Private Const API_Generic = "generic_name:"
Private Const API_AppNum = "application_number:"
Private Const API_Labeler = "labeler_name:"
Private Const API_NDC = "product_ndc:"
' API finished
Private Const API_Finished = "+AND+finished:"
' API limit/skip
Private Const API_Limit = "&limit="
Private Const API_Skip = "&skip="

Private Sub Test_NDC_API()
    Call API_NDC_PullData("rivaroxaban", 2, True, ThisWorkbook.Sheets(6))
End Sub

Public Sub API_NDC_PullData(ByVal strSearch As String, _
    ByVal searchType As Integer, _
    ByVal boolFinished As Boolean, _
    outputSheet As Worksheet)
    
    ' Warning statement
    Dim retVal As Long
    retVal = MsgBox("Warning: '" & outputSheet.Name & "' will be cleared." & vbNewLine & _
        "Are you sure you would like to proceed?", vbYesNo + vbCritical)
    If retVal = vbNo Then Exit Sub
    
    ' Search by
    '   Brand name
    '   Generic name
    '   Application number
    '   NDC (manufacturer code and product code)
    '   Labeler
    ' Choose between
    '   Finished vs unfinished products
    
    ' API variables
    Const intLimit As Integer = 100 ' Max: 100
    Dim strAPI As String
    
    On Error GoTo ErrHandling
    
    ' Create API link
    Select Case searchType
        Case 0 ' Brand name
            strAPI = API_NDC_Directory & API_Brand & strSearch & _
                API_Finished & LCase(CStr(boolFinished)) & API_Limit & intLimit
        Case 1 ' Application number
            strAPI = API_NDC_Directory & API_AppNum & strSearch & _
                API_Finished & LCase(CStr(boolFinished)) & API_Limit & intLimit
        Case 2 ' Generic name
            strAPI = API_NDC_Directory & API_Generic & strSearch & _
                API_Finished & LCase(CStr(boolFinished)) & API_Limit & intLimit
        Case 3 ' NDC
            strAPI = API_NDC_Directory & API_NDC & strSearch & _
                API_Finished & LCase(CStr(boolFinished)) & API_Limit & intLimit
        Case 4 ' Labeler
            strAPI = API_NDC_Directory & API_Labeler & strSearch & _
                API_Finished & LCase(CStr(boolFinished)) & API_Limit & intLimit
    End Select
    
    ' JSON variables
    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    Dim vFlat
    ' Loop variables
    Dim numNDC As Integer
    Dim lngTotal As Long
    Dim lngSkip As Long
    ' Worksheet headers
    Dim ndcHeader(15) As String
    ndcHeader(0) = "Brand Name"
    ndcHeader(1) = "Package NDC"
    ndcHeader(2) = "Strength"
    ndcHeader(3) = "Dosage Form"
    ndcHeader(4) = "Route"
    ndcHeader(5) = "Application Number"
    ndcHeader(6) = "Labeler Name"
    ndcHeader(7) = "Product NDC"
    ndcHeader(8) = "Generic Name"
    ndcHeader(9) = "Active Ingredients"
    ndcHeader(10) = "Product Type"
    ndcHeader(11) = "Marketing Start Date"
    ndcHeader(12) = "Listing Expiration Date"
    ndcHeader(13) = "Marketing Category"
    ndcHeader(14) = "Package Description"
    ndcHeader(15) = "Pharm Class"
    
    ' Run fast mode
    Call FastMode(True)
    
    ' Clear worksheet
    outputSheet.Cells.Delete
    
    ' Display progress bar and declare variables
    Dim pctProgress As Single
    With formProgress
        .lblProgress.Width = 0
        .lblCaption.Caption = "Pulling data..."
        .Show
    End With
    
    ' Loop API call until all data retrieved
    While lngSkip < lngTotal Or lngTotal = 0
        ' Retrieve JSON response
        With CreateObject("MSXML2.XMLHTTP")
            .Open "GET", strAPI & API_Skip & lngSkip, True
            .Send
            Do Until .ReadyState = 4: DoEvents: Loop
            sJSONString = .ResponseText
        End With
        ' Parse JSON response
        JSON.Parse sJSONString, vJSON, sState
        ' Check response validity
        Select Case True
            Case sState <> "Object"
                Unload formProgress
                MsgBox "Invalid JSON response"
                Exit Sub
            Case Not vJSON.Exists("results")
                Unload formProgress
                MsgBox "JSON contains no results"
                Exit Sub
            Case Else
                ' Convert JSON nested rows array to 2D Array and output to worksheet #1
                'Output ThisWorkbook.Sheets(1), vJSON("results")
                ' Flatten JSON
                JSON.Flatten vJSON, vFlat
                ' Convert to 2D Array and output to worksheet #2
                output ThisWorkbook.Sheets(2), vFlat
                ' Serialize JSON and save to file
                'CreateObject("Scripting.FileSystemObject") _
                    .OpenTextFile(ThisWorkbook.Path & "\sample.json", 2, True, -1) _
                    .Write JSON.Serialize(vJSON)
                ' Convert JSON to YAML and save to file
                'CreateObject("Scripting.FileSystemObject") _
                    '.OpenTextFile(ThisWorkbook.Path & "\sample.yaml", 2, True, -1) _
                    '.Write JSON.ToYaml(vJSON)
                
                ' Get total number of package NDCs and total search results
                numNDC = (Len(sJSONString) - Len(Replace(sJSONString, "package_ndc", ""))) / Len("package_ndc")
                lngTotal = vJSON("meta")("results")("total")
                ' Output extracted data to sheet
                OutputByNDC outputSheet, vJSON, numNDC
                ' Update progress bar
                pctProgress = lngSkip / lngTotal
                With formProgress
                    .lblCaption = "Pulling data: " & lngSkip & " of " & lngTotal & _
                        " (" & Int(pctProgress * 100) & "%)"
                    .lblProgress.Width = pctProgress * .frameProgress.Width
                End With
                DoEvents
        End Select
        lngSkip = lngSkip + 100
        ' Close progress bar
        If lngSkip >= lngTotal Then Unload formProgress
    Wend
       
    ' Add headers to worksheet; bold header and add filters
    With outputSheet
        .Cells(1, 1).Resize(1, UBound(ndcHeader) - LBound(ndcHeader) + 1).Font.Bold = True
        .Cells(1, 1).Resize(1, UBound(ndcHeader) - LBound(ndcHeader) + 1).AutoFilter
        .Cells(1, 1).Resize(1, UBound(ndcHeader) - LBound(ndcHeader) + 1).Value = ndcHeader
        .Columns.AutoFit
        .Activate
    End With
    
    ' Tidy up
    Dim lastRow As Long
    lastRow = outputSheet.Cells(outputSheet.Rows.Count, "B").End(xlUp).Row
    Set outputSheet = Nothing
    Call FastMode(False)
    ' Freeze pane the first row
    Range("A2").Select
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .FreezePanes = True
    End With
    ' Send message to user
    MsgBox "Completed" & vbNewLine & _
        lastRow - 1 & " row(s) of data retrieved.", vbInformation
    Exit Sub
    
ErrHandling:
    Call FastMode(False)
    If formProgress.Visible = True Then Unload formProgress
    MsgBox "Run-time error '" & Err.Number & "':" & vbNewLine & Err.Description, vbInformation
End Sub

Private Sub output(oTarget As Worksheet, vJSON)
    
    Dim aData()
    Dim aHeader()
    
    ' Convert JSON to 2D Array
    JSON.ToArray vJSON, aData, aHeader
    ' Output to target worksheet range
    With oTarget
        .Activate
        .Cells.Delete
        With .Cells(1, 1)
            .Resize(1, UBound(aHeader) - LBound(aHeader) + 1).Value = aHeader
            .Offset(1, 0).Resize( _
                    UBound(aData, 1) - LBound(aData, 1) + 1, _
                    UBound(aData, 2) - LBound(aData, 2) + 1 _
                ).Value = aData
        End With
        .Columns.AutoFit
    End With

End Sub
Private Sub OutputByNDC(ByVal outputSheet As Worksheet, _
    vJSON, ByVal arrSize As Integer)

    ' JSON variables
    Dim oItem
    Dim objPackage
    Dim objActiveIngr
    Dim strActiveIngrName As String
    Dim strActiveIngrStrength As String
    Dim objRoute
    Dim strRoute As String
    Dim objPharmClass
    Dim strPharmClass As String
    ' Worksheet variables
    Dim tempOutput() As String
    ReDim tempOutput(arrSize - 1, 0 To 15)
    Dim i As Long
    Dim lastRow As Long
    
    On Error Resume Next
    ' VBA will raise error if oItem does not contain "active_ingredients" subcategory
    ' VBA will raise error if oItem does not contain "pharm_class" subcategory
    ' Tell VBA to ignore error and process the remaining
    
    ' Loop through all items in JSON
    For Each oItem In vJSON("results")
        ' Retrieve all active ingredient names and strengths
        '   VBA will raise error if oItem does not contain "active_ingredients" subcategory
        For Each objActiveIngr In oItem("active_ingredients")
            If Len(strActiveIngrName) = 0 Then
                strActiveIngrName = objActiveIngr("name")
                strActiveIngrStrength = objActiveIngr("strength")
            Else
                strActiveIngrName = strActiveIngrName & "; " & objActiveIngr("name")
                strActiveIngrStrength = strActiveIngrStrength & ", " & objActiveIngr("strength")
            End If
        Next objActiveIngr
        ' Retrieve all routes
        '   VBA will raise error if oItem does not contain "route" subcategory
        For Each objRoute In oItem("route")
            If Len(strRoute) = 0 Then
                strRoute = objRoute
            Else
                strRoute = strRoute & ", " & objRoute
            End If
        Next objRoute
        ' Retrieve all pharm classes
        '   VBA will raise error if oItem does not contain "pharm_class" subcategory
        For Each objPharmClass In oItem("pharm_class")
            If Len(strPharmClass) = 0 Then
                strPharmClass = objPharmClass
            Else
                strPharmClass = strPharmClass & ", " & objPharmClass
            End If
        Next objPharmClass
        ' Retrieve all package NDCs and store to array
        For Each objPackage In oItem("packaging")
            tempOutput(i, 0) = oItem("brand_name")
            tempOutput(i, 1) = objPackage("package_ndc")
            tempOutput(i, 2) = strActiveIngrStrength
            tempOutput(i, 3) = oItem("dosage_form")
            tempOutput(i, 4) = strRoute
            tempOutput(i, 5) = oItem("application_number")
            tempOutput(i, 6) = oItem("labeler_name")
            tempOutput(i, 7) = oItem("product_ndc")
            tempOutput(i, 8) = oItem("generic_name")
            tempOutput(i, 9) = strActiveIngrName
            tempOutput(i, 10) = oItem("product_type")
            tempOutput(i, 11) = oItem("marketing_start_date")
            tempOutput(i, 12) = oItem("listing_expiration_date")
            tempOutput(i, 13) = oItem("marketing_category")
            tempOutput(i, 14) = objPackage("description")
            tempOutput(i, 15) = strPharmClass
            If Len(strActiveIngrStrength) = 0 Then _
                tempOutput(i, 2) = "N/A"
            If Len(tempOutput(i, 4)) = 0 Then _
                tempOutput(i, 4) = "N/A"
            If Len(strActiveIngrName) = 0 Then _
                tempOutput(i, 9) = "N/A"
            If Len(strPharmClass) = 0 Then _
                tempOutput(i, 15) = "N/A"
            i = i + 1
        Next objPackage
        ' Reset variables for next loop iteration
        strActiveIngrName = ""
        strActiveIngrStrength = ""
        strRoute = ""
        strPharmClass = ""
    Next oItem
    
    ' Output to worksheet; append to last row
    With outputSheet
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        .Cells(1 + lastRow, 1).Resize( _
            UBound(tempOutput, 1) - LBound(tempOutput, 1) + 1, _
            UBound(tempOutput, 2) - LBound(tempOutput, 2) + 1).Value = tempOutput
    End With
    
End Sub
