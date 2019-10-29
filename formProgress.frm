VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formProgress 
   Caption         =   "Progress"
   ClientHeight    =   984
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3744
   OleObjectBlob   =   "formProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
#If IsMac = False Then
    ' Hide the title bar if you're working on a Windows machine.
    Me.Height = Me.Height - 10
    modProgress.HideTitleBar Me
#End If
End Sub
