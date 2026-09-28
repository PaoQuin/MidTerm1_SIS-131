Attribute VB_Name = "modMemory"
Option Explicit

Public Sub CreateMemoryGrid()

    Dim ws As Worksheet
    Dim row As Integer
    Dim col As Integer
    Dim address As Integer

    Set ws = ThisWorkbook.Worksheets("Simulator")

    ' Clear the previous memory grid
    ws.Range("B4:Q19").Clear

    ' Create column headers: 0-F
    For col = 0 To 15
        ws.Cells(3, col + 2).value = Hex(col)
    Next col

    ' Create row headers: 0-F
    For row = 0 To 15
        ws.Cells(row + 4, 1).value = Hex(row)
    Next row

    ' Create the 256 memory positions
    For row = 0 To 15
        For col = 0 To 15
            address = row * 16 + col

            ' Every memory position starts with value 00h
            ws.Cells(row + 4, col + 2).value = "00h"

        Next col
    Next row

    ' Format the memory grid
    With ws.Range("B4:Q19")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' Format headers
    With ws.Range("A3:Q3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("A4:A19")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' Code segment: 00h-7Fh
    With ws.Range("B4:Q11")
        .Interior.Color = RGB(221, 235, 247)
    End With

    ' Data segment: 80h-FFh
    With ws.Range("B12:Q19")
        .Interior.Color = RGB(226, 239, 218)
    End With

    ' Labels
    ws.Range("S4").value = "CODE SEGMENT"
    ws.Range("S5").value = "00h - 7Fh"

    ws.Range("S7").value = "DATA SEGMENT"
    ws.Range("S8").value = "80h - FFh"

    ws.Range("S4:S8").Font.Bold = True

    ' Adjust dimensions
    ws.Columns("A:Q").ColumnWidth = 6
    ws.Columns("S:S").ColumnWidth = 18
    ws.Rows("3:19").RowHeight = 22

End Sub

    Public Function Read(ByVal address As Integer) As Integer

        Dim ws As Worksheet
        Dim row As Integer
        Dim col As Integer

        Set ws = ThisWorkbook.Worksheets("Simulator")

        row = address \ 16
        col = address Mod 16

        Read = CInt("&H" & Replace(ws.Cells(row + 4, col + 2).value, "h", ""))

    End Function

    Public Sub WriteMemory(ByVal address As Integer, ByVal value As Integer)

        Dim ws As Worksheet
        Dim row As Integer
        Dim col As Integer

        Set ws = ThisWorkbook.Worksheets("Simulator")

        row = address \ 16
        col = address Mod 16

        ws.Cells(row + 4, col + 2).value = _
            Right("0" & Hex(value), 2) & "h"
    End Sub
