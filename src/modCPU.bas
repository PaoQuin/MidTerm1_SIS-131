Attribute VB_Name = "modCPU"
Option Explicit

' CPU Registers
Public PC As Integer
Public IR As Integer
Public MAR As Integer
Public MDR As Integer
Public AX As Integer
Public BX As Integer

Public DecodedOpcode As Integer
Public Operand1 As Integer
Public Operand2 As Integer
Public DecodedInstruction As String
Public CPUHalted As Boolean

Public PendingResult As Integer
Public PendingTarget As Integer
Public PendingStore As Boolean
Public PendingMemoryStore As Boolean

Public CurrentPhase As Integer
Public CPURunning As Boolean
Public LogStep As Integer


Public Function GetRegisterValue(ByVal reg As Integer) As Integer

    If reg = 0 Then
        GetRegisterValue = AX
    ElseIf reg = 1 Then
        GetRegisterValue = BX
    End If

End Function

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


Public Sub Fetch()

    MAR = PC
    MDR = Read(MAR)
    IR = MDR
    PC = (PC + 1) Mod 256

    UpdateRegisterPanel

End Sub

Public Sub Decode()

    DecodedOpcode = IR
    Operand1 = 0
    Operand2 = 0
    DecodedInstruction = ""

    Select Case IR

        Case &H10
            DecodedInstruction = "MOV reg, imm"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H11
            DecodedInstruction = "MOV reg, reg"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H12
            DecodedInstruction = "LOAD reg, [addr]"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H13
            DecodedInstruction = "STORE [addr], reg"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H20
            DecodedInstruction = "ADD reg, imm"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H21
            DecodedInstruction = "ADD reg, reg"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H22
            DecodedInstruction = "SUB reg, imm"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H23
            DecodedInstruction = "SUB reg, reg"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H24
            DecodedInstruction = "INC reg"
            Operand1 = ReadNextByte()

        Case &H25
            DecodedInstruction = "DEC reg"
            Operand1 = ReadNextByte()

        Case &H26
            DecodedInstruction = "CMP reg, imm"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H27
            DecodedInstruction = "CMP reg, reg"
            Operand1 = ReadNextByte()
            Operand2 = ReadNextByte()

        Case &H30
            DecodedInstruction = "JMP addr"
            Operand1 = ReadNextByte()

        Case &H31
            DecodedInstruction = "JZ addr"
            Operand1 = ReadNextByte()

        Case &H32
            DecodedInstruction = "JNZ addr"
            Operand1 = ReadNextByte()

        Case &HFF
            DecodedInstruction = "HLT"

        Case Else
            DecodedInstruction = "INVALID"
            CPUHalted = True

    End Select

    UpdateRegisterPanel

End Sub

Public Function ReadNextByte() As Integer

    MAR = PC
    MDR = Read(MAR)

    ReadNextByte = MDR

    PC = (PC + 1) Mod 256

End Function


Public Sub Execute()
    Dim a As Integer
    Dim b As Integer

    PendingStore = False
    PendingMemoryStore = False
    PendingTarget = Operand1

    ' First operand is a register. Second is a value (imm) or a register.
    a = GetRegisterValue(Operand1)
    If DecodedOpcode = &H20 Or DecodedOpcode = &H22 Or DecodedOpcode = &H26 Then
        b = Operand2                        ' immediate value
    Else
        b = GetRegisterValue(Operand2)      ' register
    End If

    Select Case DecodedOpcode

        Case &H10   ' MOV reg, imm
            PendingResult = Operand2
            PendingStore = True

        Case &H11   ' MOV reg, reg
            PendingResult = b
            PendingStore = True

        Case &H12   ' LOAD reg, [addr]
            MAR = Operand2
            MDR = Read(MAR)
            PendingResult = MDR
            PendingStore = True

        Case &H13   ' STORE [addr], reg
            MAR = Operand1
            MDR = b
            PendingMemoryStore = True

        Case &H20, &H21   ' ADD
            PendingResult = ALUAdd(a, b)
            UpdateFlags PendingResult, IIf(a + b > 255, 1, 0)
            PendingStore = True

        Case &H22, &H23   ' SUB
            PendingResult = ALUSub(a, b)
            UpdateFlags PendingResult, IIf(a < b, 1, 0)
            PendingStore = True

        Case &H24   ' INC (CF stays the same)
            PendingResult = ALUInc(a)
            UpdateFlags PendingResult, CF
            PendingStore = True

        Case &H25   ' DEC (CF stays the same)
            PendingResult = ALUDec(a)
            UpdateFlags PendingResult, CF
            PendingStore = True

        Case &H26, &H27   ' CMP: only flags, no result saved
            UpdateFlags ALUCmp(a, b), IIf(a < b, 1, 0)

        Case &H30   ' JMP
            PC = Operand1

        Case &H31   ' JZ
            If ZF = 1 Then PC = Operand1

        Case &H32   ' JNZ
            If ZF = 0 Then PC = Operand1

        Case &HFF   ' HLT
            CPUHalted = True

    End Select

    UpdateRegisterPanel
End Sub

Public Sub Store()

    If PendingStore Then

        If PendingTarget = 0 Then
            AX = PendingResult
        ElseIf PendingTarget = 1 Then
            BX = PendingResult
        End If

    ElseIf PendingMemoryStore Then

        WriteMemory MAR, MDR

    End If

    PendingStore = False
    PendingMemoryStore = False

    UpdateRegisterPanel

End Sub

Public Sub StepCPU()
    If CPUHalted Then Exit Sub

    Select Case CurrentPhase

        Case 0
            Fetch
            LogMicroOperation "FETCH"
            CurrentPhase = 1

        Case 1
            Decode
            LogMicroOperation "DECODE"
            CurrentPhase = 2

        Case 2
            Execute
            LogMicroOperation "EXECUTE"
            CurrentPhase = 3

        Case 3
            Store
            LogMicroOperation "STORE"
            CurrentPhase = 0

    End Select

    ' Highlight the next active phase
    HighlightPhase CurrentPhase

    UpdateRegisterPanel
End Sub

Public Sub LogMicroOperation(ByVal phaseName As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Simulator")

    nextRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).row + 1

    ws.Cells(nextRow, "S").value = LogStep
    ws.Cells(nextRow, "T").value = phaseName

    LogStep = LogStep + 1

End Sub

Public Sub ResetCPU()

    PC = 0
    IR = 0
    MAR = 0
    MDR = 0
    AX = 0
    BX = 0

    ZF = 0
    CF = 0
    SF = 0

    DecodedOpcode = 0
    Operand1 = 0
    Operand2 = 0
    DecodedInstruction = ""

    PendingResult = 0
    PendingTarget = 0
    PendingStore = False
    PendingMemoryStore = False

    CPUHalted = False
    CurrentPhase = 0
    LogStep = 0
    
    HighlightPhase 0
    
    ThisWorkbook.Worksheets("Simulator").Range("S22:T1000").ClearContents

    UpdateRegisterPanel

End Sub

Public Sub LoadMultiplicationProgram()

    ' Clear code segment
    Dim i As Integer
    For i = 0 To 127
        WriteMemory i, 0
    Next i

    ' Multiplication program: 3 x 4 = 12
    ' 00: LOAD AX, [81h]
    WriteMemory 0, &H12
    WriteMemory 1, &H0
    WriteMemory 2, &H81

    ' 03: CMP AX, 00h
    WriteMemory 3, &H26
    WriteMemory 4, &H0
    WriteMemory 5, &H0

    ' 06: JZ 15h
    WriteMemory 6, &H31
    WriteMemory 7, &H15

    ' 08: DEC AX
    WriteMemory 8, &H25
    WriteMemory 9, &H0

    ' 0A: STORE [81h], AX
    WriteMemory 10, &H13
    WriteMemory 11, &H81
    WriteMemory 12, &H0

    ' 0D: LOAD AX, [80h]
    WriteMemory 13, &H12
    WriteMemory 14, &H0
    WriteMemory 15, &H80

    ' 10: ADD BX, AX
    WriteMemory 16, &H21
    WriteMemory 17, &H1
    WriteMemory 18, &H0

    ' 13: JMP 00h
    WriteMemory 19, &H30
    WriteMemory 20, &H0

    ' 15: HLT
    WriteMemory 21, &HFF

    ' Initial data
    ' 80h = multiplicand = 3
    ' 81h = counter = 4
    ' 82h = expected result = 12
    WriteMemory &H80, 3
    WriteMemory &H81, 4
    WriteMemory &H82, 12

    ResetCPU

End Sub

Public Sub LoadCountdownProgram()

    Dim i As Integer

    ' Clear code segment
    For i = 0 To 127
        WriteMemory i, 0
    Next i

    ' Countdown program: 5 -> 0
    ' 00: LOAD AX, [80h]
    WriteMemory 0, &H12
    WriteMemory 1, &H0
    WriteMemory 2, &H80

    ' 03: DEC AX
    WriteMemory 3, &H25
    WriteMemory 4, &H0

    ' 05: STORE [80h], AX
    WriteMemory 5, &H13
    WriteMemory 6, &H80
    WriteMemory 7, &H0

    ' 08: CMP AX, 00h
    WriteMemory 8, &H26
    WriteMemory 9, &H0
    WriteMemory 10, &H0

    ' 0B: JNZ 03h
    WriteMemory 11, &H32
    WriteMemory 12, &H3

    ' 0D: HLT
    WriteMemory 13, &HFF

    ' Initial data
    ' 80h = starting value = 5
    WriteMemory &H80, 5

    ResetCPU

End Sub


Public Sub RunCPU()

    Dim ws As Worksheet
    Dim delayMs As Double
    Dim startTime As Double

    Set ws = ThisWorkbook.Worksheets("Simulator")

    If CPUHalted Then Exit Sub

    CPURunning = True

    delayMs = ws.Range("T23").value

    If delayMs < 0 Then delayMs = 0

    Do While CPURunning And Not CPUHalted

        StepCPU

        startTime = Timer

        Do While Timer < startTime + (delayMs / 1000)
            DoEvents
        Loop

    Loop

    CPURunning = False

End Sub

Public Sub PauseCPU()

    CPURunning = False

End Sub

Public Sub HighlightPhase(ByVal phase As Integer)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Simulator")

    ' Clear previous phase highlight
    ws.Range("V10:W13").Interior.Pattern = xlNone

    Select Case phase
        Case 0
            ws.Range("V10:W10").Interior.Color = RGB(255, 255, 0)
        Case 1
            ws.Range("V11:W11").Interior.Color = RGB(255, 255, 0)
        Case 2
            ws.Range("V12:W12").Interior.Color = RGB(255, 255, 0)
        Case 3
            ws.Range("V13:W13").Interior.Color = RGB(255, 255, 0)
    End Select

End Sub

Public Sub InspectMemory()

    Dim ws As Worksheet
    Dim address As Integer
    Dim value As Integer

    Set ws = ThisWorkbook.Worksheets("Simulator")

    ' Read address entered by the user
    address = CInt("&H" & Replace(ws.Range("Y11").value, "h", ""))

    ' Validate address
    If address < 0 Or address > 255 Then
        MsgBox "Address must be between 00h and FFh."
        Exit Sub
    End If

    value = Read(address)

    ' Display address
    ws.Range("Y11").value = Right("0" & Hex(address), 2) & "h"

    ' Display value in different formats
    ws.Range("Y12").value = Right("0" & Hex(value), 2) & "h"
    ws.Range("Y13").NumberFormat = "@"
    ws.Range("Y13").value = DecToBinary(value)
    ws.Range("Y14").value = value

    ' Display mnemonic
    ws.Range("Y15").value = GetMnemonic(value)

End Sub

Public Function DecToBinary(ByVal value As Integer) As String

    Dim i As Integer
    Dim result As String

    result = ""

    For i = 7 To 0 Step -1
        If (value And (2 ^ i)) <> 0 Then
            result = result & "1"
        Else
            result = result & "0"
        End If
    Next i

    DecToBinary = result

End Function

Public Function GetMnemonic(ByVal opcode As Integer) As String

    Select Case opcode

        Case &H10: GetMnemonic = "MOV reg, imm"
        Case &H11: GetMnemonic = "MOV reg, reg"
        Case &H12: GetMnemonic = "LOAD reg, [addr]"
        Case &H13: GetMnemonic = "STORE [addr], reg"
        Case &H20: GetMnemonic = "ADD reg, imm"
        Case &H21: GetMnemonic = "ADD reg, reg"
        Case &H22: GetMnemonic = "SUB reg, imm"
        Case &H23: GetMnemonic = "SUB reg, reg"
        Case &H24: GetMnemonic = "INC reg"
        Case &H25: GetMnemonic = "DEC reg"
        Case &H26: GetMnemonic = "CMP reg, imm"
        Case &H27: GetMnemonic = "CMP reg, reg"
        Case &H30: GetMnemonic = "JMP addr"
        Case &H31: GetMnemonic = "JZ addr"
        Case &H32: GetMnemonic = "JNZ addr"
        Case &HFF: GetMnemonic = "HLT"
        Case Else: GetMnemonic = "DATA / UNKNOWN"

    End Select

End Function

