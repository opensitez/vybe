*> vybe-test: cobol/write_statement/write_basic_compiles
*> origin: languages/cobol/tests/cobol/test_write_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat".
DATA DIVISION.
FILE SECTION.
FD F.
01 R PIC X(20).
PROCEDURE DIVISION.
    OPEN OUTPUT F.
    WRITE R.
    CLOSE F.
    STOP RUN.

