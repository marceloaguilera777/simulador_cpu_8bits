Attribute VB_Name = "modALU"
' ============================================================
' modALU: Unidad Aritmetico-Logica de 8 bits
' Hace la operacion, actualiza las banderas ZF, CF, SF
' y devuelve el resultado (0 - 255).
' ============================================================
Option Explicit

Function ALU(ByVal operacion As String, ByVal a As Integer, ByVal b As Integer) As Integer
    Dim r As Integer
    Dim carry As Integer
    carry = 0

    Select Case operacion
        Case "ADD"
            r = a + b
            If r > 255 Then carry = 1      ' acarreo: no entra en 8 bits
        Case "SUB", "CMP"
            r = a - b
            If r < 0 Then carry = 1        ' prestamo: a < b
        Case "INC"
            r = a + 1
            If r > 255 Then carry = 1
        Case "DEC"
            r = a - 1
            If r < 0 Then carry = 1
        Case "AND"
            r = a And b
        Case "OR"
            r = a Or b
        Case "XOR"
            r = a Xor b
        Case "NOT"
            r = Not a
    End Select

    r = r And 255   ' el resultado se queda con 8 bits (ej: -1 -> FFh, 256 -> 00h)

    ' ---- Banderas ----
    If r = 0 Then PonerBandera "ZF", 1 Else PonerBandera "ZF", 0
    PonerBandera "CF", carry
    If r >= 128 Then PonerBandera "SF", 1 Else PonerBandera "SF", 0   ' bit 7 = 1

    ALU = r
End Function

' Prueba de la ALU con casos borde
Sub ProbarALU()
    Dim r As Integer
    Dim msg As String

    r = ALU("ADD", &HFF, 1)
    msg = msg & "FF + 01 = " & AHex(r) & Banderas() & vbCrLf
    r = ALU("SUB", 0, 1)
    msg = msg & "00 - 01 = " & AHex(r) & Banderas() & vbCrLf
    r = ALU("SUB", 5, 5)
    msg = msg & "05 - 05 = " & AHex(r) & Banderas() & vbCrLf
    r = ALU("AND", &HC, &HA)
    msg = msg & "0C AND 0A = " & AHex(r) & Banderas() & vbCrLf
    r = ALU("NOT", &HF, 0)
    msg = msg & "NOT 0F = " & AHex(r) & Banderas() & vbCrLf

    MsgBox msg, , "Prueba de la ALU"
End Sub

Function Banderas() As String
    Banderas = "    ZF=" & LeerBandera("ZF") & "  CF=" & LeerBandera("CF") & "  SF=" & LeerBandera("SF")
End Function
