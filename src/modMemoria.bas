Attribute VB_Name = "modMemoria"
' ============================================================
' modMemoria: memoria RAM de 256 bytes (00h - FFh)
' Esta en la tabla C5:R20 de la hoja CPU.
' Fila = primer digito hex de la direccion, columna = segundo digito.
' Ejemplo: la direccion 82h esta en la fila 8, columna 2.
' ============================================================
Option Explicit

' Devuelve la celda de la hoja que corresponde a una direccion
Function CeldaMemoria(ByVal direccion As Integer) As Range
    direccion = direccion And 255
    Set CeldaMemoria = Sheets("CPU").Cells(5 + direccion \ 16, 3 + direccion Mod 16)
End Function

' Read(address):  MAR <- direccion,  MDR <- RAM[MAR]
Function LeerMemoria(ByVal direccion As Integer) As Integer
    PonerRegistro "MAR", direccion
    PonerRegistro "MDR", CInt("&H" & CeldaMemoria(direccion).Value)
    LeerMemoria = LeerRegistro("MDR")
End Function

' Write(address, value):  MAR <- direccion,  MDR <- valor,  RAM[MAR] <- MDR
Sub EscribirMemoria(ByVal direccion As Integer, ByVal valor As Integer)
    PonerRegistro "MAR", direccion
    PonerRegistro "MDR", valor
    CeldaMemoria(direccion).Value = AHex(valor And 255)
End Sub

' Prueba rapida: escribe 07h en la direccion 80h y la vuelve a leer
Sub ProbarMemoria()
    EscribirMemoria &H80, 7
    MsgBox "Se escribio 07h en 80h." & vbCrLf & _
           "Leido: " & AHex(LeerMemoria(&H80)) & "h" & vbCrLf & _
           "MAR = " & AHex(LeerRegistro("MAR")) & "h   MDR = " & AHex(LeerRegistro("MDR")) & "h"
End Sub
