Attribute VB_Name = "modAPI_NDC_Dict"
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
' Dictionary object
Private objOutput As Object

Private Sub TestAPIDict()
    Call API_NDC_PullData_Dict("ceftriaxone", 2, True, ThisWorkbook.ActiveSheet)
End Sub

Public Sub API_NDC_PullData_Dict(ByVal strSearch As String, _
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
                'JSON.Flatten vJSON, vFlat
                ' Convert to 2D Array and output to worksheet #2
                'output ThisWorkbook.Sheets(2), vFlat
                ' Serialize JSON and save to file
                'CreateObject("Scripting.FileSystemObject") _
                    .OpenTextFile(ThisWorkbook.Path & "\sample.json", 2, True, -1) _
                    .Write JSON.Serialize(vJSON)
                ' Convert JSON to YAML and save to file
                'CreateObject("Scripting.FileSystemObject") _
                    '.OpenTextFile(ThisWorkbook.Path & "\sample.yaml", 2, True, -1) _
                    '.Write JSON.ToYaml(vJSON)
                
                ' Get total number of package NDCs and total search results
                'numNDC = (Len(sJSONString) - Len(Replace(sJSONString, "package_ndc", ""))) / Len("package_ndc")
                lngTotal = vJSON("meta")("results")("total")
                ' Flatten extracted data to dictionary
                Call FlattenToDict(vJSON)
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
    
    ' Write to worksheet and clear objOutput
    Call WriteToSheet(objOutput, outputSheet)
    Set objOutput = Nothing
    
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

Public Sub FlattenToDict(vJSON)
    
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
    ' Dictionary variables
    Dim dictNDC As New clsNDC
    
    On Error Resume Next
    ' VBA will raise error if oItem does not contain "active_ingredients" subcategory
    ' VBA will raise error if oItem does not contain "pharm_class" subcategory
    ' Tell VBA to ignore error and process the remaining
    
    ' Initialize dictionary
    Call InitializeDict(objOutput)
    
    ' Loop through all items in JSON
    For Each oItem In vJSON("results")
        ' Concatenate active ingredient names and strengths
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
        ' Concatenate routes
        '   VBA will raise error if oItem does not contain "route" subcategory
        For Each objRoute In oItem("route")
            If Len(strRoute) = 0 Then
                strRoute = objRoute
            Else
                strRoute = strRoute & ", " & objRoute
            End If
        Next objRoute
        ' Concatenate pharm classes
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
            Set dictNDC = New clsNDC
            dictNDC.BrandName = oItem("brand_name")
            dictNDC.PkgNDC = objPackage("package_ndc")
            dictNDC.ActiveIngrStrength = strActiveIngrStrength
            dictNDC.DosageForm = oItem("dosage_form")
            dictNDC.Route = strRoute
            dictNDC.AppNum = oItem("application_number")
            dictNDC.Mfr = oItem("labeler_name")
            dictNDC.ProdNDC = oItem("product_ndc")
            dictNDC.GenericName = oItem("generic_name")
            dictNDC.ActiveIngrName = strActiveIngrName
            dictNDC.ProdType = oItem("product_type")
            dictNDC.MktgStartDate = oItem("marketing_start_date")
            dictNDC.ListingExpDate = oItem("listing_expiration_date")
            dictNDC.MktgCategory = oItem("marketing_category")
            dictNDC.PkgDescription = objPackage("description")
            dictNDC.PharmClass = strPharmClass
            Call AddItemDict(objOutput, dictNDC.PkgNDC, dictNDC)
        Next objPackage
        ' Reset variables for next loop iteration
        strActiveIngrName = ""
        strActiveIngrStrength = ""
        strRoute = ""
        strPharmClass = ""
    Next oItem

    ' Clear class
    Set dictNDC = Nothing

End Sub

Public Sub InitializeDict(objDict As Object)

    ' Create dictionary object (erases current dictionary)
    Set objDict = CreateObject("Scripting.Dictionary")
    objDict.CompareMode = vbTextCompare 'Not case-sensitive
    
End Sub

Public Sub AddItemDict(objDict As Object, strNDC As String, _
    data, Optional boolOverwrite As Boolean = False)
    
    ' Check if dictionary exists
    If ExistsDict(objDict) = False Then
        MsgBox "Dictionary does not exist", vbInformation, "AddItemDict()"
        Exit Sub
    End If
    ' Add item to dictionary if the key doesn't exist or boolOverwrite = True
    If objDict.Exists(strNDC) = True And boolOverwrite = False Then
        Exit Sub
    Else
        objDict.Add strNDC, data
    End If
    
End Sub

Public Function ExistsDict(objDict As Object) As Boolean

    If objDict Is Nothing Then
        ExistsDict = False
    Else
        ExistsDict = True
    End If
    
End Function

Public Sub WriteToImmediate(objDict As Object)
    
    ' Validate dictionary
    If objDict.Count = 0 Then
        MsgBox "Dictionary count is 0", vbInformation, "WriteToImmediate()"
        Exit Sub
    End If
    
    ' Declare variables
    Dim key As Variant, oNDC As New clsNDC
    
    ' Read through the dictionary
    For Each key In objDict.keys
        Set oNDC = objDict(key)
        With oNDC
            ' Write to the Immediate Window (Ctrl + G)
            Debug.Print .BrandName, .PkgNDC, .ActiveIngrStrength, .DosageForm, _
                .Route, .AppNum, .Mfr, .ProdNDC, _
                .GenericName, .ActiveIngrName, .ProdType, .MktgStartDate, _
                .ListingExpDate, .MktgCategory, .PkgDescription, .PharmClass
        End With
    Next key
    
    ' Clear class
    Set oNDC = Nothing
    
End Sub

Public Sub WriteToSheet(objDict As Object, sheet As Worksheet)
    
    ' Validate dictionary
    If objDict.Count = 0 Then
        MsgBox "Dictionary count is 0", vbInformation, "WriteToSheet()"
        Exit Sub
    End If
    
    ' Declare variables
    Dim key As Variant, oNDC As New clsNDC
    Dim i As Long, arrData() As String
    
    ' Resize array
    ReDim arrData(objDict.Count - 1, 0 To 15)
    
    ' Read through the dictionary
    i = 0
    For Each key In objDict.keys
        Set oNDC = objDict(key)
        With oNDC
            ' Copy to the array
            arrData(i, 0) = .BrandName
            arrData(i, 1) = .PkgNDC
            arrData(i, 2) = .ActiveIngrStrength
            arrData(i, 3) = .DosageForm
            arrData(i, 4) = .Route
            arrData(i, 5) = .AppNum
            arrData(i, 6) = .Mfr
            arrData(i, 7) = .ProdNDC
            arrData(i, 8) = .GenericName
            arrData(i, 9) = .ActiveIngrName
            arrData(i, 10) = .ProdType
            arrData(i, 11) = .MktgStartDate
            arrData(i, 12) = .ListingExpDate
            arrData(i, 13) = .MktgCategory
            arrData(i, 14) = .PkgDescription
            arrData(i, 15) = .PharmClass
            i = i + 1
        End With
    Next key
    
    ' Write to the worksheet
    sheet.Cells(2, 1).Resize( _
        UBound(arrData, 1) - LBound(arrData, 1) + 1, _
        UBound(arrData, 2) - LBound(arrData, 2) + 1).Value = arrData
    
    ' Clear class
    Set oNDC = Nothing
    
End Sub
