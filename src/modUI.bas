Attribute VB_Name = "modUI"
Option Explicit

' Formats a byte as two hex digits, for example 5 -> "05h"
Public Function HexByte(ByVal value As Integer) As String
    HexByte = Right("0" & Hex(value), 2) & "h"
End Function

' Adds one line to the micro-operation log (columns V:W, from row 18)
Public Sub LogMicroOperation(ByVal phaseName As String, ByVal detail As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Simulator")
    nextRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).row + 1
    If nextRow < 18 Then nextRow = 18

    ws.Cells(nextRow, "V").value = LogStep
    ws.Cells(nextRow, "W").value = "[Step " & Format(LogStep, "00") & "] " & phaseName & ": " & detail

    LogStep = LogStep + 1
End Sub
