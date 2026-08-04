*> vybe-test: cobol/open_close_modes/open_input_output_mixed_compiles
*> origin: languages/cobol/tests/cobol/test_open_close_modes.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F1 ASSIGN TO "a.dat".
    SELECT F2 ASSIGN TO "b.dat".
DATA DIVISION.
FILE SECTION.
FD F1.
01 R1 PIC X(10).
FD F2.
01 R2 PIC X(10).
PROCEDURE DIVISION.
    OPEN INPUT F1.
    OPEN OUTPUT F2.
    CLOSE F1 F2.
    STOP RUN.

