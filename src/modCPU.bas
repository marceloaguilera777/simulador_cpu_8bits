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
    Dim fase As String
    fase = Sheets("CPU").Range("U24").Value

    If fase = "DETENIDO" Then
        MsgBox "El CPU esta detenido (HLT). Presione RESET para empezar de nuevo."
        Exit Sub
    End If

    QuitarResaltado     ' se borra el naranja de la fase anterior
    Select Case fase
        Case "", "STORE"
            Fetch
            fase = "FETCH"
        Case "FETCH"
            Decode
            fase = "DECODE"
        Case "DECODE"
            Execute
            fase = "EXECUTE"
        Case "EXECUTE"
            Store
            fase = "STORE"
    End Select
    PonerFase fase
    EscribirLog fase, DetalleFase(fase)

    If fase = "DECODE" And operacion = "???" Then
        MsgBox "Opcode desconocido: " & AHex(LeerRegistro("IR")) & "h. El CPU se detiene."
        PonerFase "DETENIDO"
    End If
    If fase = "EXECUTE" And operacion = "HLT" Then PonerFase "DETENIDO"
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
    BorrarLog
    QuitarResaltado
End Sub

' Muestra la fase y le pone un color distinto a cada una
Sub PonerFase(ByVal fase As String)
    With Sheets("CPU").Range("U24")
        .Value = fase
        Select Case fase
            Case "FETCH":    .Interior.Color = RGB(155, 194, 230)   ' azul
            Case "DECODE":   .Interior.Color = RGB(204, 192, 218)   ' lila
            Case "EXECUTE":  .Interior.Color = RGB(255, 192, 0)     ' naranja
            Case "STORE":    .Interior.Color = RGB(169, 208, 142)   ' verde
            Case "DETENIDO": .Interior.Color = RGB(255, 124, 128)   ' rojo
            Case Else:       .Interior.ColorIndex = xlNone
        End Select
    End With
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

' Texto para el log de lo que paso en cada fase
Function DetalleFase(ByVal fase As String) As String
    Dim mar As String, mdr As String
    mar = AHex(LeerRegistro("MAR"))
    mdr = AHex(LeerRegistro("MDR"))

    Select Case fase
        Case "FETCH"
            DetalleFase = "MAR=" & mar & ", MDR=" & mdr & " -> IR=" & AHex(LeerRegistro("IR")) & _
                          ", PC=" & AHex(LeerRegistro("PC"))

        Case "DECODE"
            DetalleFase = "IR=" & AHex(LeerRegistro("IR")) & " -> " & TextoInstruccion()
            If dosBytes Then DetalleFase = DetalleFase & "  (operando: MAR=" & mar & ", MDR=" & mdr & ")"

        Case "EXECUTE"
            Select Case operacion
                Case "ADD", "SUB", "CMP", "AND", "OR", "XOR", "INC", "DEC", "NOT"
                    DetalleFase = "ALU " & operacion & " = " & AHex(resultado) & Banderas()
                Case "MOV"
                    DetalleFase = "dato = " & AHex(resultado)
                Case "LOAD"
                    DetalleFase = "MAR=" & mar & ", MDR=" & mdr & " (leido de memoria)"
                Case "STORE"
                    DetalleFase = "dato a guardar = " & AHex(resultado)
                Case "JMP"
                    DetalleFase = "PC=" & AHex(operando) & " (salta)"
                Case "JZ"
                    If LeerBandera("ZF") = 1 Then
                        DetalleFase = "ZF=1 -> PC=" & AHex(operando) & " (salta)"
                    Else
                        DetalleFase = "ZF=0 -> no salta"
                    End If
                Case "JNZ"
                    If LeerBandera("ZF") = 0 Then
                        DetalleFase = "ZF=0 -> PC=" & AHex(operando) & " (salta)"
                    Else
                        DetalleFase = "ZF=1 -> no salta"
                    End If
                Case "HLT"
                    DetalleFase = "HLT -> se detiene el reloj"
            End Select

        Case "STORE"
            If Not guardar Then
                DetalleFase = "(no guarda nada)"
            ElseIf operacion = "STORE" Then
                DetalleFase = "MAR=" & mar & ", MDR=" & mdr & " -> RAM[" & mar & "]=" & mdr
            Else
                DetalleFase = registro & " <- " & AHex(resultado)
            End If
    End Select
End Function
