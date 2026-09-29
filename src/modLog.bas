Attribute VB_Name = "modLog"
' ============================================================
' modLog: resaltado de lo que esta activo + log de micro-operaciones
' - Lo que cambia en cada fase se pinta de naranja
' - El log esta debajo de la memoria (desde B24), el paso mas nuevo arriba
' ============================================================
Option Explicit

Dim numeroPaso As Integer

' Pinta una celda de naranja (componente activo en esta fase)
Sub Resaltar(ByVal celda As Range)
    celda.Interior.Color = RGB(255, 192, 0)
End Sub

' Vuelve a poner los colores normales
Sub QuitarResaltado()
    With Sheets("CPU")
        .Range("C5:R12").Interior.Color = RGB(221, 235, 247)   ' codigo
        .Range("C13:R20").Interior.Color = RGB(226, 239, 218)  ' datos
        .Range("U5:U10").Interior.Color = RGB(255, 242, 204)   ' registros
        .Range("U13:U15").Interior.Color = RGB(255, 242, 204)  ' banderas
    End With
End Sub

' Agrega una linea al log, ej: [Paso 01] FETCH: MAR=00, MDR=10 -> IR=10, PC=01
Sub EscribirLog(ByVal fase As String, ByVal detalle As String)
    With Sheets("CPU")
        numeroPaso = numeroPaso + 1
        .Range("B25:B400").Value = .Range("B24:B399").Value    ' baja las lineas anteriores
        .Range("B24").Value = "[Paso " & Format(numeroPaso, "00") & "] " & fase & ": " & detalle
    End With
End Sub

' Borra el log (se usa en RESET)
Sub BorrarLog()
    numeroPaso = 0
    With Sheets("CPU")
        .Range("B24:B400").ClearContents
        .Range("B23").Value = "LOG DE MICRO-OPERACIONES (el paso mas nuevo arriba)"
        .Range("B23").Font.Bold = True
    End With
End Sub
