# MidTerm1_SIS-131
Simulador de CPU von Neumann de 8 bits con memoria de 256 bytes (Excel + VBA).

## Instruction Set Architecture (ISA)

### Register Encoding

| Code | Register |
|---|---|
| 00h | AX |
| 01h | BX |

### Instruction Set

| Opcode | Mnemonic | Bytes | Description | Flags |
|---|---|---:|---|---|
| 10h | MOV reg, imm | 3 | Moves an immediate value into AX or BX | None |
| 11h | MOV reg, reg | 3 | Copies one register into another | None |
| 12h | LOAD reg, [addr] | 3 | Loads a byte from memory into a register | None |
| 13h | STORE [addr], reg | 3 | Stores a register value into memory | None |
| 20h | ADD reg, imm | 3 | Adds an immediate value to a register | ZF, CF, SF |
| 21h | ADD reg, reg | 3 | Adds one register to another | ZF, CF, SF |
| 22h | SUB reg, imm | 3 | Subtracts an immediate value from a register | ZF, CF, SF |
| 23h | SUB reg, reg | 3 | Subtracts one register from another | ZF, CF, SF |
| 24h | INC reg | 2 | Increments a register by 1 | ZF, SF; CF unchanged |
| 25h | DEC reg | 2 | Decrements a register by 1 | ZF, SF; CF unchanged |
| 26h | CMP reg, imm | 3 | Compares a register with an immediate value | ZF, CF, SF |
| 27h | CMP reg, reg | 3 | Compares two registers | ZF, CF, SF |
| 30h | JMP addr | 2 | Unconditional jump to an address | None |
| 31h | JZ addr | 2 | Jumps if ZF = 1 | None |
| 32h | JNZ addr | 2 | Jumps if ZF = 0 | None |
| FFh | HLT | 1 | Stops CPU execution | None |

### Flag Behavior

- **ZF (Zero Flag):** 1 when the ALU result is 00h.
- **CF (Carry Flag):** 1 when an unsigned addition overflows or a subtraction/compare requires a borrow.
- **SF (Sign Flag):** Bit 7 of the 8-bit result.
- `MOV`, `LOAD`, `STORE`, jumps and `HLT` do not modify flags.


## CPU Architecture

```mermaid
flowchart LR
    MEM["Memory<br/>256 × 8-bit"] <-->|"Address / Data Bus"| CPU["CPU"]

    CPU --> CU["Control Unit"]
    CPU --> REG["Registers<br/>PC • IR • MAR • MDR • AX • BX"]
    CPU --> ALU["ALU<br/>ADD • SUB • INC • DEC<br/>AND • OR • XOR • NOT • CMP"]

    CU --> REG
    CU --> ALU
    REG <--> ALU

## Program Trace

### Multiplication Program: 3 × 4 = 12

The demo program performs multiplication by repeated addition.

Initial data:

| Address | Value | Meaning |
|---|---:|---|
| 80h | 03h | Multiplicand |
| 81h | 04h | Counter |
| 82h | 0Ch | Expected result |

Initial registers:

| Register | Value |
|---|---|
| PC | 00h |
| AX | 00h |
| BX | 00h |

### Execution Flow

1. `LOAD AX, [81h]` loads the counter into AX.
2. `CMP AX, 00h` checks whether the counter reached zero.
3. `JZ 15h` exits the loop when ZF = 1.
4. `DEC AX` decreases the counter.
5. `STORE [81h], AX` updates the counter in memory.
6. `LOAD AX, [80h]` loads the multiplicand.
7. `ADD BX, AX` adds the multiplicand to the accumulated result.
8. `JMP 00h` repeats the loop.
9. When the counter reaches zero, `JZ` jumps to `HLT`.

### Register Trace

| Iteration | Counter | AX | BX |
|---:|---:|---:|---:|
| Initial | 04h | 00h | 00h |
| 1 | 03h | 03h | 03h |
| 2 | 02h | 03h | 06h |
| 3 | 01h | 03h | 09h |
| 4 | 00h | 03h | 0Ch |
| Final | 00h | 03h | 0Ch |

Final result:

`BX = 0Ch = 12 decimal`

The CPU then reaches `HLT` and stops execution.

The micro-operation log records the execution chronologically as:

FETCH → DECODE → EXECUTE → STORE

for each CPU cycle until the program reaches HLT.
