Attribute VB_Name = "modCPU"
' ============================================================
' modCPU: Unidad de Control
' Ciclo de instruccion en 4 fases: FETCH -> DECODE -> EXECUTE -> STORE
' Cada llamada a Paso() ejecuta UNA fase.
' La fase que se acaba de ejecutar se muestra en la celda U24.
' ============================================================
Option Explicit

' Datos de la instruccion actual (los llena DECODE)
Dim operacion As String    ' "MOV", "ADD", "JZ", ...
Dim registro As String     ' registro destino: "AX" o "BX"
Dim dosBytes As Boolean    ' True si la instruccion trae un segundo byte
Dim operando As Integer    ' segundo byte: numero inmediato o direccion
Dim resultado As Integer   ' lo calcula EXECUTE
Dim guardar As Boolean     ' True si STORE tiene que guardar algo

' ---------- Ejecuta la siguiente fase del ciclo ----------
Sub Paso()
    Select Case Sheets("CPU").Range("U24").Value
        Case "", "STORE"
            Fetch
            PonerFase "FETCH"
        Case "FETCH"
            Decode
            PonerFase "DECODE"
            If operacion = "???" Then
                MsgBox "Opcode desconocido: " & AHex(LeerRegistro("IR")) & "h. El CPU se detiene."
                PonerFase "DETENIDO"
            End If
        Case "DECODE"
            Execute
            If operacion = "HLT" Then PonerFase "DETENIDO" Else PonerFase "EXECUTE"
        Case "EXECUTE"
            Store
            PonerFase "STORE"
        Case "DETENIDO"
            MsgBox "El CPU esta detenido (HLT). Use Reiniciar para empezar de nuevo."
    End Select
End Sub

' ---------- Ejecuta todo el programa hasta HLT ----------
Sub Ejecutar()
    Dim pasos As Integer
    Do While Sheets("CPU").Range("U24").Value <> "DETENIDO" And pasos < 2000
        Paso
        pasos = pasos + 1
        DoEvents
    Loop
End Sub

' ---------- RESET: registros en cero y ciclo desde FETCH ----------
Sub Reiniciar()
    ResetCPU
    PonerFase ""
    Sheets("CPU").Range("U25").Value = ""
End Sub

Sub PonerFase(ByVal fase As String)
    Sheets("CPU").Range("U24").Value = fase
End Sub

' ============================================================
' FASE 1 - FETCH: MAR <- PC, MDR <- RAM[MAR], IR <- MDR, PC <- PC + 1
' ============================================================
Sub Fetch()
    Dim pc As Integer
    pc = LeerRegistro("PC")
    PonerRegistro "IR", LeerMemoria(pc)     ' LeerMemoria usa MAR y MDR
    PonerRegistro "PC", pc + 1
End Sub

' ============================================================
' FASE 2 - DECODE: interpreta el opcode que esta en IR
'   primer digito hex  = operacion   (IR \ 16)
'   segundo digito hex = modo        (IR Mod 16)
' ============================================================
Sub Decode()
    Dim op As Integer, modo As Integer
    op = LeerRegistro("IR") \ 16
    modo = LeerRegistro("IR") Mod 16

    ' Registro: modo par (0, 2) = AX, modo impar (1, 3) = BX
    If modo Mod 2 = 0 Then registro = "AX" Else registro = "BX"

    Select Case op
        Case 0: operacion = "HLT"
        Case 1: operacion = "MOV"
        Case 2: If modo <= 1 Then operacion = "LOAD" Else operacion = "STORE"
        Case 3: operacion = "ADD"
        Case 4: operacion = "SUB"
        Case 5: If modo <= 1 Then operacion = "INC" Else operacion = "DEC"
        Case 6: operacion = "CMP"
        Case 7: operacion = "AND"
        Case 8: operacion = "OR"
        Case 9: operacion = "XOR"
        Case 10: operacion = "NOT"
        Case 15
            If modo = 0 Then operacion = "JMP"
            If modo = 1 Then operacion = "JZ"
            If modo = 2 Then operacion = "JNZ"
            If modo > 2 Then operacion = "???"
        Case Else: operacion = "???"
    End Select

    ' Instrucciones de 2 bytes: LOAD/STORE, saltos y las que usan numero inmediato
    dosBytes = False
    Select Case op
        Case 2, 15
            dosBytes = True
        Case 1, 3, 4, 6, 7, 8, 9
            If modo <= 1 Then dosBytes = True
    End Select

    ' Si tiene 2 bytes, se lee el operando: MAR <- PC, MDR <- RAM[MAR], PC <- PC + 1
    operando = 0
    If dosBytes Then
        operando = LeerMemoria(LeerRegistro("PC"))
        PonerRegistro "PC", LeerRegistro("PC") + 1
    End If

    Sheets("CPU").Range("U25").Value = TextoInstruccion()
End Sub

' ============================================================
' FASE 3 - EXECUTE: la ALU calcula o se resuelve el salto
' ============================================================
Sub Execute()
    guardar = False
    Select Case operacion
        Case "MOV"
            resultado = ValorFuente()
            guardar = True
        Case "LOAD"
            resultado = LeerMemoria(operando)            ' MAR <- dir, MDR <- RAM[dir]
            guardar = True
        Case "STORE"
            resultado = LeerRegistro(registro)
            guardar = True
        Case "ADD", "SUB", "AND", "OR", "XOR"
            resultado = ALU(operacion, LeerRegistro(registro), ValorFuente())
            guardar = True
        Case "CMP"
            resultado = ALU("CMP", LeerRegistro(registro), ValorFuente())   ' solo banderas
        Case "INC", "DEC", "NOT"
            resultado = ALU(operacion, LeerRegistro(registro), 0)
            guardar = True
        Case "JMP"
            PonerRegistro "PC", operando
        Case "JZ"
            If LeerBandera("ZF") = 1 Then PonerRegistro "PC", operando
        Case "JNZ"
            If LeerBandera("ZF") = 0 Then PonerRegistro "PC", operando
        Case "HLT"
            ' no hace nada: el ciclo se detiene
    End Select
End Sub

' ============================================================
' FASE 4 - STORE: guarda el resultado en el registro o en memoria
' ============================================================
Sub Store()
    If Not guardar Then Exit Sub     ' CMP y saltos no guardan nada
    If operacion = "STORE" Then
        EscribirMemoria operando, resultado      ' MDR -> RAM[MAR]
    Else
        PonerRegistro registro, resultado
    End If
End Sub

' ---------- Funciones de ayuda ----------

' Segundo operando: el numero inmediato o el otro registro
Function ValorFuente() As Integer
    If dosBytes Then
        ValorFuente = operando
    ElseIf registro = "AX" Then
        ValorFuente = LeerRegistro("BX")
    Else
        ValorFuente = LeerRegistro("AX")
    End If
End Function

' Texto de la instruccion, ej: "ADD AX, 05" o "STORE [82], AX"
Function TextoInstruccion() As String
    Dim otro As String
    If registro = "AX" Then otro = "BX" Else otro = "AX"

    Select Case operacion
        Case "HLT"
            TextoInstruccion = "HLT"
        Case "LOAD"
            TextoInstruccion = "LOAD " & registro & ", [" & AHex(operando) & "]"
        Case "STORE"
            TextoInstruccion = "STORE [" & AHex(operando) & "], " & registro
        Case "INC", "DEC", "NOT"
            TextoInstruccion = operacion & " " & registro
        Case "JMP", "JZ", "JNZ"
            TextoInstruccion = operacion & " " & AHex(operando)
        Case "???"
            TextoInstruccion = "???"
        Case Else
            If dosBytes Then
                TextoInstruccion = operacion & " " & registro & ", " & AHex(operando)
            Else
                TextoInstruccion = operacion & " " & registro & ", " & otro
            End If
    End Select
End Function
