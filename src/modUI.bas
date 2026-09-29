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

' Removes the previous highlight from registers, flags and memory
Public Sub ClearHighlights()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulator")

    ws.Range("T11:T16").Interior.Pattern = xlNone
    ws.Range("T19:T21").Interior.Pattern = xlNone
    ws.Range("B4:Q11").Interior.Color = RGB(221, 235, 247)   ' code segment
    ws.Range("B12:Q19").Interior.Color = RGB(226, 239, 218)  ' data segment
End Sub

' Highlights one register of the panel (row 11 = PC ... row 16 = BX)
Public Sub HighlightRegister(ByVal row As Integer)
    ThisWorkbook.Worksheets("Simulator").Cells(row, "T").Interior.Color = RGB(255, 230, 153)
End Sub

' Highlights the memory cell at an address
Public Sub HighlightMemory(ByVal address As Integer)
    ThisWorkbook.Worksheets("Simulator").Cells(address \ 16 + 4, address Mod 16 + 2).Interior.Color = RGB(248, 203, 173)
End Sub
