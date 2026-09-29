' ============================================================
' Codigo de la hoja CPU (se pega dentro de la hoja, no se importa)
' ============================================================
Option Explicit

' Al hacer clic en una celda de la memoria, se muestra en el inspector
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim celda As Range
    Set celda = Target.Cells(1, 1)

    If Not Intersect(celda, Range("C5:R20")) Is Nothing Then
        Range("U18").Value = AHex((celda.Row - 5) * 16 + (celda.Column - 3))
        Range("U19").Value = celda.Value
    End If
End Sub
