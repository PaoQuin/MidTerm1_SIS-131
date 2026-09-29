Attribute VB_Name = "modCPU"
Option Explicit

' CPU Registers
Public PC As Integer
Public IR As Integer
Public MAR As Integer
Public MDR As Integer
Public AX As Integer
Public BX As Integer

Public Sub CreateRegisterPanel()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Simulator")

    ws.Range("S10:T21").Clear

    ws.Range("S10").value = "CPU REGISTERS"

    ws.Range("S11").value = "PC"
    ws.Range("S12").value = "IR"
    ws.Range("S13").value = "MAR"
    ws.Range("S14").value = "MDR"
    ws.Range("S15").value = "AX"
    ws.Range("S16").value = "BX"

    ws.Range("S18").value = "FLAGS"

    ws.Range("S19").value = "ZF"
    ws.Range("S20").value = "CF"
    ws.Range("S21").value = "SF"

    ws.Range("T11:T16").value = "00h"
    ws.Range("T19:T21").value = 0

    With ws.Range("S10:T21")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("S10:T10").Merge
    ws.Range("S18:T18").Merge

    ws.Range("S10").Font.Bold = True
    ws.Range("S18").Font.Bold = True

    ws.Columns("S:T").ColumnWidth = 12

End Sub

Public Sub UpdateRegisterPanel()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Simulator")

    ws.Range("T11").value = Right("0" & Hex(PC), 2) & "h"
    ws.Range("T12").value = Right("0" & Hex(IR), 2) & "h"
    ws.Range("T13").value = Right("0" & Hex(MAR), 2) & "h"
    ws.Range("T14").value = Right("0" & Hex(MDR), 2) & "h"
    ws.Range("T15").value = Right("0" & Hex(AX), 2) & "h"
    ws.Range("T16").value = Right("0" & Hex(BX), 2) & "h"

    ws.Range("T19").value = ZF
    ws.Range("T20").value = CF
    ws.Range("T21").value = SF

End Sub
