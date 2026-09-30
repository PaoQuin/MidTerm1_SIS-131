Attribute VB_Name = "modStyle"
Option Explicit

' TEMPORARY: run once to style the sheet, then remove this module
Public Sub FormatSheetOnce()
    Dim ws As Worksheet
    Dim i As Integer
    Set ws = ThisWorkbook.Worksheets("Simulator")
    ws.Activate
    Application.DisplayAlerts = False

    ' Base font and gridlines
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ActiveWindow.DisplayGridlines = False

    ' Column widths
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B:Q").ColumnWidth = 6
    ws.Columns("R").ColumnWidth = 2
    ws.Columns("S:T").ColumnWidth = 13
    ws.Columns("U").ColumnWidth = 2
    ws.Columns("V").ColumnWidth = 14
    ws.Columns("W").ColumnWidth = 58
    ws.Columns("X").ColumnWidth = 13
    ws.Columns("Y").ColumnWidth = 18
    ws.Range("Z:XFD").EntireColumn.Hidden = True

    ' Title bar
    StyleHeader ws.Range("A1:Y1"), "8-BIT CPU SIMULATOR  |  SIS-131 Computer Architecture"
    ws.Range("A1").Font.Size = 16
    ws.Rows(1).RowHeight = 34

    ' Memory
    StyleHeader ws.Range("B2:Q2"), "MAIN MEMORY (256 BYTES)"
    With ws.Range("A3:Q3,A4:A19")
        .Interior.Color = RGB(47, 84, 150)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B4:Q19").Font.Name = "Consolas"
    ws.Rows("4:19").RowHeight = 20
    LightBorders ws.Range("B4:Q19")

    ' Memory map legend
    StyleHeader ws.Range("S2:T2"), "MEMORY MAP"
    ws.Range("S4:T4").Merge
    ws.Range("S5:T5").Merge
    ws.Range("S7:T7").Merge
    ws.Range("S8:T8").Merge
    ws.Range("S4:T5").Interior.Color = RGB(221, 235, 247)
    ws.Range("S7:T8").Interior.Color = RGB(226, 239, 218)
    ws.Range("S4,S7").Font.Bold = True
    ws.Range("S4:T8").HorizontalAlignment = xlCenter
    LightBorders ws.Range("S4:T5")
    LightBorders ws.Range("S7:T8")

    ' Registers and flags
    StyleHeader ws.Range("S10:T10"), ""
    StyleHeader ws.Range("S18:T18"), ""
    ws.Range("S11:S16,S19:S21").Font.Bold = True
    ws.Range("S11:S16,S19:S21").Interior.Color = RGB(242, 242, 242)
    ws.Range("T11:T21").Font.Name = "Consolas"
    ws.Range("T11:T21").Font.Size = 11
    LightBorders ws.Range("S10:T16")
    LightBorders ws.Range("S18:T21")

    ' Instruction cycle phases
    StyleHeader ws.Range("V9:W9"), "INSTRUCTION CYCLE"
    For i = 10 To 13
        ws.Range("V" & i & ":W" & i).Merge
    Next i
    ws.Range("V10:W13").HorizontalAlignment = xlCenter
    ws.Range("V10:W13").Font.Bold = True
    LightBorders ws.Range("V10:W13")

    ' Delay for RUN
    ws.Range("V15").value = "DELAY (ms)"
    ws.Range("V15").Font.Bold = True
    If ws.Range("W15").value = "" Then ws.Range("W15").value = 300
    ws.Range("W15").Interior.Color = RGB(255, 242, 204)
    ws.Range("W15").HorizontalAlignment = xlLeft
    LightBorders ws.Range("V15:W15")

    ' Micro-operation log
    StyleHeader ws.Range("V17:W17"), "MICRO-OPERATION LOG"
    ws.Range("V18:V1000").HorizontalAlignment = xlCenter
    ws.Range("V18:W1000").Font.Name = "Consolas"
    ws.Range("V18:W1000").Font.Size = 9

    ' Memory inspector
    StyleHeader ws.Range("X10:Y10"), ""
    ws.Range("X11:X15").Font.Bold = True
    ws.Range("X11:X15").Interior.Color = RGB(242, 242, 242)
    ws.Range("Y11:Y15").HorizontalAlignment = xlCenter
    ws.Range("Y11:Y13").Font.Name = "Consolas"
    ws.Range("Y11").Interior.Color = RGB(255, 242, 204)
    LightBorders ws.Range("X11:Y15")

    ' Buttons: remove the old ones and build two clean rows
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).OnAction <> "" Then ws.Shapes(i).Delete
    Next i
    ws.Rows(20).RowHeight = 20
    ws.Rows(21).RowHeight = 24
    ws.Rows(22).RowHeight = 20
    ws.Rows(23).RowHeight = 24
    AddButton ws, "STEP", "StepCPU", ws.Range("B21:E21"), RGB(31, 56, 100)
    AddButton ws, "RUN", "RunCPU", ws.Range("F21:I21"), RGB(31, 56, 100)
    AddButton ws, "PAUSE", "PauseCPU", ws.Range("J21:M21"), RGB(31, 56, 100)
    AddButton ws, "RESET", "ResetCPU", ws.Range("N21:Q21"), RGB(31, 56, 100)
    AddButton ws, "LOAD PROGRAM", "LoadMultiplicationProgram", ws.Range("B23:E23"), RGB(122, 135, 153)
    AddButton ws, "COUNTDOWN", "LoadCountdownProgram", ws.Range("F23:I23"), RGB(122, 135, 153)
    AddButton ws, "INSPECT", "InspectMemory", ws.Range("J23:M23"), RGB(122, 135, 153)

    ' View
    ws.Range("A1:Y24").Select
    ActiveWindow.Zoom = True
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    Application.DisplayAlerts = True
End Sub

Private Sub StyleHeader(ByVal rng As Range, ByVal text As String)
    rng.Merge
    If text <> "" Then rng.Cells(1, 1).value = text
    With rng
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub LightBorders(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal area As Range, ByVal fillColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 3, area.Top + 2, area.Width - 6, area.Height - 4)
    shp.Line.Visible = msoFalse
    shp.Fill.ForeColor.RGB = fillColor
    shp.OnAction = macroName
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.text = caption
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub


' Extra styling: full register names, quick guide and dark console log
Public Sub StyleExtras()
    Dim ws As Worksheet
    Dim i As Integer
    Set ws = ThisWorkbook.Worksheets("Simulator")

    ' Small gaps
    ws.Range("A2").Interior.Color = RGB(31, 56, 100)
    ws.Range("S17:T17").Borders.LineStyle = xlNone
    ws.Range("S16:T16").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("S16:T16").Borders(xlEdgeBottom).Color = RGB(191, 191, 191)

    ' Full register and flag names (labels only, values in column T are untouched)
    ws.Columns("S").ColumnWidth = 24
    ws.Range("S11").value = "PC (Program Counter)"
    ws.Range("S12").value = "IR (Instruction Reg)"
    ws.Range("S13").value = "MAR (Mem Address Reg)"
    ws.Range("S14").value = "MDR (Mem Data Reg)"
    ws.Range("S15").value = "AX (Accumulator)"
    ws.Range("S16").value = "BX (General)"
    ws.Range("S19").value = "ZF (Zero Flag)"
    ws.Range("S20").value = "CF (Carry Flag)"
    ws.Range("S21").value = "SF (Sign Flag)"
    ws.Range("S11:S16,S19:S21").HorizontalAlignment = xlLeft
    ws.Range("S11:S16,S19:S21").IndentLevel = 1
    ws.Range("T11:T21").Font.Color = RGB(31, 56, 100)
    ws.Range("T11:T21").Font.Bold = True

    ' Quick guide in the empty top-right area
    StyleHeader ws.Range("V2:Y2"), "QUICK GUIDE"
    ws.Range("V3").value = "1.  LOAD PROGRAM or COUNTDOWN writes a program into memory"
    ws.Range("V4").value = "2.  STEP runs one phase;  RUN runs until HLT (delay in W15)"
    ws.Range("V5").value = "3.  PAUSE stops RUN;  RESET clears registers, flags and the log"
    ws.Range("V6").value = "4.  Type an address in Y11 and press INSPECT"
    ws.Range("V7").value = "Yellow = active phase and register   |   Orange = memory cell in use"
    For i = 3 To 7
        ws.Range("V" & i & ":Y" & i).Merge
        ws.Range("V" & i).HorizontalAlignment = xlLeft
        ws.Range("V" & i).IndentLevel = 1
    Next i
    ws.Range("V3:Y7").Interior.Color = RGB(242, 242, 242)
    ws.Range("V7").Font.Italic = True
    ws.Range("V7").Font.Color = RGB(91, 107, 127)
    LightBorders ws.Range("V3:Y7")

    ' Dark console log
    With ws.Range("V18:W400")
        .Interior.Color = RGB(30, 30, 30)
        .Font.Color = RGB(146, 208, 80)
    End With
    ws.Range("V17:W17").Interior.Color = RGB(0, 0, 0)
End Sub

