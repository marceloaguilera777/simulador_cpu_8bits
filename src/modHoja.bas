Attribute VB_Name = "modHoja"
' ============================================================
' modHoja: arma la hoja CPU (memoria, registros e inspector)
' ============================================================
Option Explicit

Sub PrepararHoja()
    Dim ws As Worksheet
    Dim i As Integer, fila As Integer
    Set ws = Sheets("CPU")

    ws.Cells.Clear
    ws.Range("A1").Value = "SIMULADOR DE CPU DE 8 BITS"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True

    ' ---------- MEMORIA RAM 16 x 16 (C5:R20) ----------
    ws.Range("C3").Value = "MEMORIA RAM (256 bytes: 00h - FFh)"
    ws.Range("C3").Font.Bold = True
    ws.Range("B4:R20").NumberFormat = "@"   ' todo como texto para que "00" no se vuelva 0

    For i = 0 To 15
        ws.Cells(4, 3 + i).Value = Hex(i)          ' columnas: 0 ... F
        ws.Cells(5 + i, 2).Value = Hex(i) & "0"    ' filas: 00, 10 ... F0
    Next i
    ws.Range("B4:R4").Font.Bold = True
    ws.Range("B5:B20").Font.Bold = True

    With ws.Range("C5:R20")
        .Value = "00"
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("C5:R12").Interior.Color = RGB(221, 235, 247)    ' codigo: 00h - 7Fh (azul)
    ws.Range("C13:R20").Interior.Color = RGB(226, 239, 218)   ' datos:  80h - FFh (verde)
    ws.Range("A5").Value = "CODIGO"
    ws.Range("A13").Value = "DATOS"

    ' ---------- REGISTROS (T5:W10) ----------
    ws.Range("T3").Value = "REGISTROS"
    ws.Range("T3").Font.Bold = True
    ws.Range("T4:W4").Value = Array("Reg", "Hex", "Binario", "Dec")
    ws.Range("T4:W4").Font.Bold = True
    ws.Range("T5").Value = "PC"
    ws.Range("T6").Value = "IR"
    ws.Range("T7").Value = "MAR"
    ws.Range("T8").Value = "MDR"
    ws.Range("T9").Value = "AX"
    ws.Range("T10").Value = "BX"
    ws.Range("U5:U15").NumberFormat = "@"

    ' Binario y decimal se calculan solos con formulas de Excel
    For fila = 5 To 10
        ws.Cells(fila, 22).Formula = "=HEX2BIN(U" & fila & ",8)"
        ws.Cells(fila, 23).Formula = "=HEX2DEC(U" & fila & ")"
    Next fila
    ws.Range("T5:W10").Interior.Color = RGB(255, 242, 204)
    ws.Range("T5:W10").Borders.LineStyle = xlContinuous

    ' ---------- BANDERAS (T13:U15) ----------
    ws.Range("T12").Value = "BANDERAS"
    ws.Range("T12").Font.Bold = True
    ws.Range("T13").Value = "ZF"
    ws.Range("T14").Value = "CF"
    ws.Range("T15").Value = "SF"
    ws.Range("T13:U15").Interior.Color = RGB(255, 242, 204)
    ws.Range("T13:U15").Borders.LineStyle = xlContinuous

    ' ---------- INSPECTOR DE MEMORIA (T18:U21) ----------
    ws.Range("T17").Value = "INSPECTOR (clic en una celda de la RAM)"
    ws.Range("T17").Font.Bold = True
    ws.Range("T18").Value = "Direccion"
    ws.Range("T19").Value = "Hex"
    ws.Range("T20").Value = "Binario"
    ws.Range("T21").Value = "Decimal"
    ws.Range("U18:U19").NumberFormat = "@"
    ws.Range("U18").Value = "00"
    ws.Range("U19").Value = "00"
    ws.Range("U20").Formula = "=HEX2BIN(U19,8)"
    ws.Range("U21").Formula = "=HEX2DEC(U19)"
    ws.Range("T18:U21").Borders.LineStyle = xlContinuous

    ' ---------- CICLO DE INSTRUCCION (T23:U25) ----------
    ws.Range("T23").Value = "CICLO DE INSTRUCCION"
    ws.Range("T23").Font.Bold = True
    ws.Range("T24").Value = "Fase"
    ws.Range("T25").Value = "Instruccion"
    ws.Range("T24:U25").Borders.LineStyle = xlContinuous
    ws.Range("U24:U25").Font.Bold = True

    ' ---------- Tamanos de columnas ----------
    ws.Range("C:R").ColumnWidth = 4
    ws.Range("T:T").ColumnWidth = 10
    ws.Range("V:V").ColumnWidth = 10

    ResetCPU
End Sub
