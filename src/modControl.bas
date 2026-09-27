Attribute VB_Name = "modControl"
Option Explicit

Public Sub ExportModules()
    Dim comp As Object
    Dim exportPath As String
    exportPath = ThisWorkbook.path & Application.PathSeparator & "src" & Application.PathSeparator

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1   ' standard module
                comp.Export exportPath & comp.Name & ".bas"
            Case 100 ' sheet or workbook code
                If comp.CodeModule.CountOfLines > 0 Then
                    comp.Export exportPath & comp.Name & ".cls"
                End If
        End Select
    Next comp
End Sub
