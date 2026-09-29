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
