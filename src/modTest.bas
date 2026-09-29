Attribute VB_Name = "modTest"
Option Explicit

Public Sub TestALU()

    Debug.Print "===== ALU TEST ====="

    ' ADD
    If ALUAdd(10, 5) = 15 Then
        Debug.Print "ADD: PASS"
    Else
        Debug.Print "ADD: FAIL"
    End If

    ' SUB
    If ALUSub(10, 5) = 5 Then
        Debug.Print "SUB: PASS"
    Else
        Debug.Print "SUB: FAIL"
    End If

    ' INC
    If ALUInc(10) = 11 Then
        Debug.Print "INC: PASS"
    Else
        Debug.Print "INC: FAIL"
    End If

    ' DEC
    If ALUDec(10) = 9 Then
        Debug.Print "DEC: PASS"
    Else
        Debug.Print "DEC: FAIL"
    End If

    ' AND
    If ALUAnd(12, 10) = 8 Then
        Debug.Print "AND: PASS"
    Else
        Debug.Print "AND: FAIL"
    End If

    ' OR
    If ALUOr(12, 10) = 14 Then
        Debug.Print "OR: PASS"
    Else
        Debug.Print "OR: FAIL"
    End If

    ' XOR
    If ALUXor(12, 10) = 6 Then
        Debug.Print "XOR: PASS"
    Else
        Debug.Print "XOR: FAIL"
    End If

    ' NOT
    If ALUNot(0) = 255 Then
        Debug.Print "NOT: PASS"
    Else
        Debug.Print "NOT: FAIL"
    End If

    ' CMP
    If ALUCmp(10, 5) = 5 Then
        Debug.Print "CMP: PASS"
    Else
        Debug.Print "CMP: FAIL"
    End If

    Debug.Print "===== FLAG TEST ====="

    UpdateFlags 0, 0

    If ZF = 1 Then
        Debug.Print "ZERO: PASS"
    Else
        Debug.Print "ZERO: FAIL"
    End If

    UpdateFlags 255, 1

    If CF = 1 Then
        Debug.Print "CARRY: PASS"
    Else
        Debug.Print "CARRY: FAIL"
    End If

    UpdateFlags 255, 0

    If SF = 1 Then
        Debug.Print "SIGN: PASS"
    Else
        Debug.Print "SIGN: FAIL"
    End If

    Debug.Print "===== TEST COMPLETE ====="

End Sub
