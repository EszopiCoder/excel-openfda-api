VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenFDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum QueryName
    VetADE
    DrugADE
    DrugLabel
    DrugNDC
    DrugRecallEnforcement
    Device510K
    DeviceClass
    DeviceRecallEnforcement
    DeviceADE
    DeviceApproval
    DeviceRecall
    DeviceListing
    DeviceID
    FoodRecallEnforcement
    FoodADE
    NDSE
    SubstanceData
End Enum
Public API_Key As String
Public API_Search As String
Public API_Count As String
Public API_Sort As String
Public API_Limit As String
Public API_Skip As String

Private pQuery As QueryName

Public Property Let API_Query(Value As QueryName)
    pQuery = Value
End Property

Public Property Get API_Link() As String
    ' Validate API_Search and API_Count
    If Len(API_Search) = 0 And Len(API_Count) = 0 Then
        API_Link = "Invalid arguments"
        Exit Sub
    End If
    ' Select FDA query
    Select Case pQuery
        Case QueryName.VetADE
            API_Link = "https://api.fda.gov/animalandveterinary/event.json?"
        Case QueryName.DrugADE
            API_Link = "https://api.fda.gov/drug/event.json?"
        Case QueryName.DrugLabel
            API_Link = "https://api.fda.gov/drug/label.json?"
        Case QueryName.DrugNDC
            API_Link = "https://api.fda.gov/drug/ndc.json?"
        Case QueryName.DrugRecallEnforcement
            API_Link = "https://api.fda.gov/drug/enforcement.json?"
        Case QueryName.Device510K
            API_Link = "https://api.fda.gov/device/510k.json?"
        Case QueryName.DeviceClass
            API_Link = "https://api.fda.gov/device/classification.json?"
        Case QueryName.DeviceRecallEnforcement
            API_Link = "https://api.fda.gov/device/enforcement.json?"
        Case QueryName.DeviceADE
            API_Link = "https://api.fda.gov/device/event.json?"
        Case QueryName.DeviceApproval
            API_Link = "https://api.fda.gov/device/pma.json?"
        Case QueryName.DeviceRecall
            API_Link = "https://api.fda.gov/device/recall.json?"
        Case QueryName.DeviceListing
            API_Link = "https://api.fda.gov/device/registrationlisting.json?"
        Case QueryName.DeviceID
            API_Link = "https://api.fda.gov/device/udi.json?"
        Case QueryName.FoodRecallEnforcement
            API_Link = "https://api.fda.gov/food/enforcement.json?"
        Case QueryName.FoodADE
            API_Link = "https://api.fda.gov/food/event.json?"
        Case QueryName.NDSE
            API_Link = "https://api.fda.gov/other/nsde.json?"
        Case QueryName.SubstanceData
            API_Link = "https://api.fda.gov/other/substance.json?"
    End Select
    ' Add API key (not necessary to run API)
    If Len(API_Key) > 0 Then _
        API_Link = API_Link & "api_key=" & API_Key
    ' Add search
    If Len(API_Search) > 0 Then _
        API_Link = API_Link & "&search=" & API_Search
    ' Add count
    If Len(API_Count) > 0 Then _
        API_Link = API_Link & "&count=" & API_Count
    ' Add sort
    If Len(API_Sort) > 0 Then _
        API_Link = API_Link & "&sort=" & API_Sort
    ' Add limit if value is greater than 0
    If Len(API_Limit) > 0 And Val(API_Limit) > 0 Then _
        API_Link = API_Link & "&limit=" & API_Limit
    ' Add skip if value is greater than 0
    If Len(API_Skip) > 0 And Val(API_Skip) > 0 Then _
        API_Link = API_Link & "&skip=" & API_Skip
    ' Fix API_Link if key is not used
    API_Link = Replace(API_Link, "?&", "?")
End Property

