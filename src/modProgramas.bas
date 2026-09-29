Attribute VB_Name = "modProgramas"
' ============================================================
' modProgramas: programas que se cargan en la memoria
' ============================================================
Option Explicit

' Escribe una lista de bytes en memoria desde una direccion
' Ejemplo: CargarBytes 0, "10 03 11 00"
Sub CargarBytes(ByVal inicio As Integer, ByVal bytes As String)
    Dim partes() As String
    Dim i As Integer
    partes = Split(bytes, " ")
    For i = 0 To UBound(partes)
        CeldaMemoria(inicio + i).Value = partes(i)
    Next i
End Sub

' Pone toda la memoria en 00
Sub BorrarMemoria()
    Sheets("CPU").Range("C5:R20").Value = "00"
End Sub

' ------------------------------------------------------------
' Programa de prueba de saltos (JNZ, JMP, JZ) y HLT
'   00: MOV AX, 03
'   02: MOV BX, 00
'   04: INC BX          <- bucle
'   05: DEC AX
'   06: JNZ 04          (repite mientras AX no sea 0)
'   08: STORE [80], BX  (guarda 03 en 80h)
'   0A: JMP 0E
'   0C: MOV BX, FF      (nunca se ejecuta)
'   0E: JZ 12           (ZF = 1, entonces salta)
'   10: MOV BX, EE      (nunca se ejecuta)
'   12: HLT
' Resultado esperado: AX = 00, BX = 03, RAM[80] = 03, ZF = 1
' ------------------------------------------------------------
Sub CargarProgramaPrueba()
    BorrarMemoria
    CargarBytes &H0, "10 03 11 00 51 52 F2 04 23 80 F0 0E 11 FF F1 12 11 EE 00"
    Reiniciar
End Sub
