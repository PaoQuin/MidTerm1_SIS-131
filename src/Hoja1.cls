VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim address As Integer
    Dim value As Integer
    Dim textValue As String

    ' Only handle changes inside the memory grid
    If Intersect(Target, Me.Range("B4:Q19")) Is Nothing Then Exit Sub

    ' Only process one memory cell at a time
    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo InvalidValue

    textValue = Trim(CStr(Target.value))
    textValue = Replace(textValue, "h", "")
    textValue = Replace(textValue, "H", "")

    ' Reject empty values
    If textValue = "" Then GoTo InvalidValue

    ' Convert hexadecimal value
    value = CInt("&H" & textValue)

    ' Accept only 00h to FFh
    If value < 0 Or value > 255 Then GoTo InvalidValue

    ' Calculate memory address from grid position
    address = (Target.row - 4) * 16 + (Target.Column - 2)

    ' Synchronize worksheet cell with RAM
    Application.EnableEvents = False
    WriteMemory address, value
    Application.EnableEvents = True

    Exit Sub

InvalidValue:
    Application.EnableEvents = False
    Application.Undo
    Application.EnableEvents = True

    MsgBox "Invalid memory value. Enter a hexadecimal value from 00h to FFh."

End Sub
