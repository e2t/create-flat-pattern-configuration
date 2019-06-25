Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Public swDoc As ModelDoc2

Sub Main()
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
    If swDoc Is Nothing Then Exit Sub
    If swDoc.GetType <> swDocPART Then Exit Sub
    GetListConf swDoc
    MainForm.Show
End Sub

Sub CreateFPforSelected(docname As String)
    Dim i As Integer
    Dim conf As String
    Dim errors As swActivateDocError_e
    Dim newDoc As ModelDoc2
    Dim drawing As DrawingDoc
    
    Set newDoc = swApp.NewDocument(swApp.GetUserPreferenceStringValue(swDefaultTemplateDrawing), _
                                   swDwgPaperA4size, 0, 0)
    Set drawing = newDoc
    For i = 0 To MainForm.lstConfNames.ListCount - 1
        If MainForm.lstConfNames.Selected(i) Then
            conf = MainForm.lstConfNames.List(i)
            drawing.CreateFlatPatternViewFromModelView3 docname, conf, 0, 0, 0, False, False
            MainForm.lstConfNames.Selected(i) = False
            MainForm.Repaint
        End If
    Next
    swApp.CloseDoc newDoc.GetPathName
    Set newDoc = Nothing
    Set drawing = Nothing
    swApp.ActivateDoc3 docname, False, swDontRebuildActiveDoc, errors
End Sub

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function

Sub GetListConf(doc As ModelDoc2)
    Dim x As Variant
    Dim conf As String
    Dim curConf As String
    
    curConf = doc.GetActiveConfiguration.Name
    For Each x In doc.GetConfigurationNames
        conf = x
        MainForm.lstConfNames.AddItem conf
        If conf = curConf Then
            MainForm.lstConfNames.Selected(MainForm.lstConfNames.ListCount - 1) = True
        End If
    Next
End Sub
