VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Создать развертки"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnCancel_Click()
    ExitApp
End Sub

Private Sub btnOk_Click()
    CreateFPforSelected swDoc.GetPathName
    ExitApp
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Integer
    
    For i = 0 To lstConfNames.ListCount - 1
        lstConfNames.Selected(i) = Not lstConfNames.List(i) Like "*SM-FLAT-PATTERN"
    Next
End Sub
