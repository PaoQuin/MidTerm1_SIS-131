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
        ws.Cells(3, col + 2).Value = Hex(col)
    Next col

    ' Create row headers: 0-F
    For row = 0 To 15
        ws.Cells(row + 4, 1).Value = Hex(row)
    Next row

    ' Create the 256 memory positions
    For row = 0 To 15
        For col = 0 To 15
            address = row * 16 + col

            ws.Cells(row + 4, col + 2).Value = _
                Right("0" & Hex(address), 2) & "h"

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
    ws.Range("S4").Value = "CODE SEGMENT"
    ws.Range("S5").Value = "00h - 7Fh"

    ws.Range("S7").Value = "DATA SEGMENT"
    ws.Range("S8").Value = "80h - FFh"

    ws.Range("S4:S8").Font.Bold = True

    ' Adjust dimensions
    ws.Columns("A:Q").ColumnWidth = 6
    ws.Columns("S:S").ColumnWidth = 18
    ws.Rows("3:19").RowHeight = 22

End Sub
