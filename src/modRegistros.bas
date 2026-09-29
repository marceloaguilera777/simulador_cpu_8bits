Attribute VB_Name = "modRegistros"
' ============================================================
' modRegistros: registros del CPU y banderas
' Cada registro esta en una celda de la columna U (en hexadecimal)
' ============================================================
Option Explicit

' Devuelve la celda donde esta guardado cada registro o bandera
Function CeldaRegistro(ByVal nombre As String) As Range
    Dim fila As Integer
    Select Case nombre
        Case "PC":  fila = 5
        Case "IR":  fila = 6
        Case "MAR": fila = 7
        Case "MDR": fila = 8
        Case "AX":  fila = 9
        Case "BX":  fila = 10
        Case "ZF":  fila = 13
        Case "CF":  fila = 14
        Case "SF":  fila = 15
    End Select
    Set CeldaRegistro = Sheets("CPU").Cells(fila, 21)   ' columna U
End Function

Function LeerRegistro(ByVal nombre As String) As Integer
    LeerRegistro = CInt("&H" & CeldaRegistro(nombre).Value)
End Function

' Guarda un valor en un registro (se queda solo con 8 bits)
Sub PonerRegistro(ByVal nombre As String, ByVal valor As Integer)
    CeldaRegistro(nombre).Value = AHex(valor And 255)
    Resaltar CeldaRegistro(nombre)
End Sub

' Banderas: solo 0 o 1
Function LeerBandera(ByVal nombre As String) As Integer
    LeerBandera = CInt(CeldaRegistro(nombre).Value)
End Function

Sub PonerBandera(ByVal nombre As String, ByVal valor As Integer)
    CeldaRegistro(nombre).Value = CStr(valor)
    Resaltar CeldaRegistro(nombre)
End Sub

' RESET: todo en cero
Sub ResetCPU()
    PonerRegistro "PC", 0
    PonerRegistro "IR", 0
    PonerRegistro "MAR", 0
    PonerRegistro "MDR", 0
    PonerRegistro "AX", 0
    PonerRegistro "BX", 0
    PonerBandera "ZF", 0
    PonerBandera "CF", 0
    PonerBandera "SF", 0
End Sub

' Convierte un numero a texto hexadecimal de 2 digitos (10 -> "0A")
Function AHex(ByVal valor As Integer) As String
    AHex = Right("0" & Hex(valor), 2)
End Function
