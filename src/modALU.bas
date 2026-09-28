Attribute VB_Name = "modALU"
Option Explicit
Public ZF As Integer
Public CF As Integer
Public SF As Integer

Public Function ALUAdd(ByVal a As Integer, ByVal b As Integer) As Integer
    ALUAdd = (a + b) Mod 256
End Function


Public Function ALUSub(ByVal a As Integer, ByVal b As Integer) As Integer
    ALUSub = (a - b) Mod 256

    If ALUSub < 0 Then
        ALUSub = ALUSub + 256
    End If
End Function


Public Function ALUInc(ByVal a As Integer) As Integer
    ALUInc = (a + 1) Mod 256
End Function


Public Function ALUDec(ByVal a As Integer) As Integer
    ALUDec = (a - 1) Mod 256

    If ALUDec < 0 Then
        ALUDec = ALUDec + 256
    End If
End Function


Public Function ALUAnd(ByVal a As Integer, ByVal b As Integer) As Integer
    ALUAnd = a And b
End Function


Public Function ALUOr(ByVal a As Integer, ByVal b As Integer) As Integer
    ALUOr = a Or b
End Function


Public Function ALUXor(ByVal a As Integer, ByVal b As Integer) As Integer
    ALUXor = a Xor b
End Function


Public Function ALUNot(ByVal a As Integer) As Integer
    ALUNot = (Not a) And 255
End Function

Public Function ALUCmp(ByVal a As Integer, ByVal b As Integer) As Integer
    ALUCmp = (a - b) Mod 256

    If ALUCmp < 0 Then
        ALUCmp = ALUCmp + 256
    End If
End Function

Public Sub UpdateFlags(ByVal result As Integer, ByVal carry As Integer)

    ' Zero Flag
    If result = 0 Then
        ZF = 1
    Else
        ZF = 0
    End If

    ' Carry Flag
    CF = carry

    ' Sign Flag
    If result >= 128 Then
        SF = 1
    Else
        SF = 0
    End If

End Sub
