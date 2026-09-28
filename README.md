# Simulador de CPU de 8 bits (von Neumann)

## Arquitectura (diagrama Mermaid)

## Mapa de memoria

La memoria principal tiene **256 posiciones de 8 bits** (00h – FFh), divididas en dos segmentos lógicos:

| Rango       | Segmento | Uso                                          |
|-------------|----------|----------------------------------------------|
| `00h – 7Fh` | Código   | Instrucciones del programa (128 bytes)       |
| `80h – FFh` | Datos    | Variables, arreglos y resultados (128 bytes) |

## Tabla ISA

### Formato del opcode

Cada opcode ocupa **1 byte** (2 dígitos hexadecimales):

- **Nibble alto** → operación (1 = MOV, 3 = ADD, 4 = SUB, ...).
- **Nibble bajo** → modo / registros involucrados.

| Nibble bajo | Modo      | Operandos | Bytes |
|-------------|-----------|-----------|-------|
| `0`         | Inmediato | `AX, imm` | 2     |
| `1`         | Inmediato | `BX, imm` | 2     |
| `2`         | Registro  | `AX, BX`  | 1     |
| `3`         | Registro  | `BX, AX`  | 1     |

La Unidad de Control decodifica con: `operación = opcode \ 16` y `modo = opcode Mod 16`.

### Repertorio de instrucciones

| Mnemónico         | Opcode | Bytes | Descripción                        | Banderas     |
|-------------------|--------|-------|------------------------------------|--------------|
| `HLT`             | `00`   | 1     | Detiene el reloj del CPU           | —            |
| `MOV AX, imm`     | `10`   | 2     | AX ← imm                           | —            |
| `MOV BX, imm`     | `11`   | 2     | BX ← imm                           | —            |
| `MOV AX, BX`      | `12`   | 1     | AX ← BX                            | —            |
| `MOV BX, AX`      | `13`   | 1     | BX ← AX                            | —            |
| `LOAD AX, [dir]`  | `20`   | 2     | AX ← RAM[dir]                      | —            |
| `LOAD BX, [dir]`  | `21`   | 2     | BX ← RAM[dir]                      | —            |
| `STORE [dir], AX` | `22`   | 2     | RAM[dir] ← AX                      | —            |
| `STORE [dir], BX` | `23`   | 2     | RAM[dir] ← BX                      | —            |
| `ADD AX, imm`     | `30`   | 2     | AX ← AX + imm                      | ZF, CF, SF   |
| `ADD BX, imm`     | `31`   | 2     | BX ← BX + imm                      | ZF, CF, SF   |
| `ADD AX, BX`      | `32`   | 1     | AX ← AX + BX                       | ZF, CF, SF   |
| `ADD BX, AX`      | `33`   | 1     | BX ← BX + AX                       | ZF, CF, SF   |
| `SUB AX, imm`     | `40`   | 2     | AX ← AX − imm                      | ZF, CF, SF   |
| `SUB BX, imm`     | `41`   | 2     | BX ← BX − imm                      | ZF, CF, SF   |
| `SUB AX, BX`      | `42`   | 1     | AX ← AX − BX                       | ZF, CF, SF   |
| `SUB BX, AX`      | `43`   | 1     | BX ← BX − AX                       | ZF, CF, SF   |
| `INC AX`          | `50`   | 1     | AX ← AX + 1                        | ZF, CF, SF   |
| `INC BX`          | `51`   | 1     | BX ← BX + 1                        | ZF, CF, SF   |
| `DEC AX`          | `52`   | 1     | AX ← AX − 1                        | ZF, CF, SF   |
| `DEC BX`          | `53`   | 1     | BX ← BX − 1                        | ZF, CF, SF   |
| `CMP AX, imm`     | `60`   | 2     | AX − imm (solo actualiza banderas) | ZF, CF, SF   |
| `CMP BX, imm`     | `61`   | 2     | BX − imm (solo actualiza banderas) | ZF, CF, SF   |
| `CMP AX, BX`      | `62`   | 1     | AX − BX (solo actualiza banderas)  | ZF, CF, SF   |
| `CMP BX, AX`      | `63`   | 1     | BX − AX (solo actualiza banderas)  | ZF, CF, SF   |
| `AND AX, imm`     | `70`   | 2     | AX ← AX AND imm                    | ZF, SF, CF=0 |
| `AND BX, imm`     | `71`   | 2     | BX ← BX AND imm                    | ZF, SF, CF=0 |
| `AND AX, BX`      | `72`   | 1     | AX ← AX AND BX                     | ZF, SF, CF=0 |
| `AND BX, AX`      | `73`   | 1     | BX ← BX AND AX                     | ZF, SF, CF=0 |
| `OR AX, imm`      | `80`   | 2     | AX ← AX OR imm                     | ZF, SF, CF=0 |
| `OR BX, imm`      | `81`   | 2     | BX ← BX OR imm                     | ZF, SF, CF=0 |
| `OR AX, BX`       | `82`   | 1     | AX ← AX OR BX                      | ZF, SF, CF=0 |
| `OR BX, AX`       | `83`   | 1     | BX ← BX OR AX                      | ZF, SF, CF=0 |
| `XOR AX, imm`     | `90`   | 2     | AX ← AX XOR imm                    | ZF, SF, CF=0 |
| `XOR BX, imm`     | `91`   | 2     | BX ← BX XOR imm                    | ZF, SF, CF=0 |
| `XOR AX, BX`      | `92`   | 1     | AX ← AX XOR BX                     | ZF, SF, CF=0 |
| `XOR BX, AX`      | `93`   | 1     | BX ← BX XOR AX                     | ZF, SF, CF=0 |
| `NOT AX`          | `A0`   | 1     | AX ← NOT AX                        | ZF, SF       |
| `NOT BX`          | `A1`   | 1     | BX ← NOT BX                        | ZF, SF       |
| `JMP dir`         | `F0`   | 2     | PC ← dir                           | —            |
| `JZ dir`          | `F1`   | 2     | Si ZF = 1 → PC ← dir               | —            |
| `JNZ dir`         | `F2`   | 2     | Si ZF = 0 → PC ← dir               | —            |

### Banderas (registro de estado)

| Bandera | Nombre     | Se activa (1) cuando...                                                  |
|---------|------------|--------------------------------------------------------------------------|
| `ZF`    | Zero Flag  | El resultado de la última operación es `00h`                             |
| `CF`    | Carry Flag | Hay acarreo en la suma (resultado > FFh) o préstamo en la resta (a < b)  |
| `SF`    | Sign Flag  | El bit más significativo (bit 7) del resultado es 1                      |

> **Decisión de diseño:** `HLT` se codifica como `00h`. Si el programa se sale hacia memoria vacía, el CPU se detiene en lugar de ejecutar datos basura.

## Manual de usuario

## Programa demostrativo y traza