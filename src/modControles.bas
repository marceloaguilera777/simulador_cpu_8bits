Attribute VB_Name = "modControles"
' ============================================================
' modControles: botones STEP, RUN, PAUSE, RESET y LOAD PROGRAM
' El retardo del modo RUN se cambia en la celda U2 (milisegundos)
' ============================================================
Option Explicit

Dim pausado As Boolean

' Crea los botones en la fila 2 de la hoja
Sub CrearBotones()
    Dim ws As Worksheet
    Set ws = Sheets("CPU")
    ws.Buttons.Delete
    ws.Rows(2).RowHeight = 26

    AgregarBoton ws, 0, 70, "STEP", "BotonStep"
    AgregarBoton ws, 75, 70, "RUN", "BotonRun"
    AgregarBoton ws, 150, 70, "PAUSE", "BotonPausa"
    AgregarBoton ws, 225, 70, "RESET", "BotonReset"
    AgregarBoton ws, 300, 110, "LOAD PROGRAM", "BotonLoad"

    ws.Range("T2").Value = "Delay (ms)"
    ws.Range("T2").Font.Bold = True
    ws.Range("U2").Value = 300
    ws.Range("U2").Interior.Color = RGB(255, 242, 204)
    ws.Range("U2").Borders.LineStyle = xlContinuous
End Sub

Sub AgregarBoton(ByVal ws As Worksheet, ByVal x As Integer, ByVal ancho As Integer, _
                 ByVal texto As String, ByVal macro As String)
    Dim b As Object
    Set b = ws.Buttons.Add(ws.Range("C2").Left + x, ws.Range("C2").Top + 2, ancho, 22)
    b.Caption = texto
    b.OnAction = macro
End Sub

' STEP: ejecuta una sola fase del ciclo
Sub BotonStep()
    Paso
End Sub

' RUN: ejecuta fase por fase con el delay de U2 hasta HLT o PAUSE
Sub BotonRun()
    pausado = False
    Do While Sheets("CPU").Range("U24").Value <> "DETENIDO" And Not pausado
        Paso
        Esperar Val(Sheets("CPU").Range("U2").Value)
    Loop
End Sub

' PAUSE: detiene el RUN (se puede seguir con STEP o RUN)
Sub BotonPausa()
    pausado = True
End Sub

' RESET: registros y PC a cero (la memoria no se borra)
Sub BotonReset()
    pausado = True
    Reiniciar
End Sub

' LOAD PROGRAM: carga el programa demo (multiplicacion 7 x 5)
Sub BotonLoad()
    pausado = True
    CargarProgramaDemo
End Sub

' Espera unos milisegundos sin congelar Excel (asi se puede apretar PAUSE)
Sub Esperar(ByVal ms As Double)
    Dim inicio As Single
    inicio = Timer
    Do While Timer < inicio + ms / 1000
        DoEvents
    Loop
    DoEvents
End Sub
