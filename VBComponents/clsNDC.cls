VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BrandName As String
Public PkgNDC As String
Private pActiveIngrStrength As String
Public DosageForm As String
Private pRoute As String
Public AppNum As String
Public Mfr As String
Public ProdNDC As String
Public GenericName As String
Private pActiveIngrName As String
Public ProdType As String
Public MktgStartDate As String
Public ListingExpDate As String
Public MktgCategory As String
Public PkgDescription As String
Private pPharmClass As String

Private Sub Class_Initialize()
    BrandName = "N/A"
    PkgNDC = "N/A"
    pActiveIngrStrength = "N/A"
    DosageForm = "N/A"
    pRoute = "N/A"
    AppNum = "N/A"
    Mfr = "N/A"
    ProdNDC = "N/A"
    GenericName = "N/A"
    pActiveIngrName = "N/A"
    ProdType = "N/A"
    MktgStartDate = "N/A"
    ListingExpDate = "N/A"
    MktgCategory = "N/A"
    PkgDescription = "N/A"
    pPharmClass = "N/A"
End Sub

Public Property Let ActiveIngrStrength(strText As String)
    If Len(strText) = 0 Then
        pActiveIngrStrength = "N/A"
    Else
        pActiveIngrStrength = strText
    End If
End Property
Public Property Get ActiveIngrStrength() As String
    ActiveIngrStrength = pActiveIngrStrength
End Property

Public Property Let Route(strText As String)
    If Len(strText) = 0 Then
        pRoute = "N/A"
    Else
        pRoute = strText
    End If
End Property
Public Property Get Route() As String
    Route = pRoute
End Property

Public Property Let ActiveIngrName(strText As String)
    If Len(strText) = 0 Then
        pActiveIngrName = "N/A"
    Else
        pActiveIngrName = strText
    End If
End Property
Public Property Get ActiveIngrName() As String
    ActiveIngrName = pActiveIngrName
End Property

Public Property Let PharmClass(strText As String)
    If Len(strText) = 0 Then
        pPharmClass = "N/A"
    Else
        pPharmClass = strText
    End If
End Property
Public Property Get PharmClass() As String
    PharmClass = pPharmClass
End Property
