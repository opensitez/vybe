*> vybe-test: cobol/write_statement/write_line_sequential_compiles
*> origin: languages/cobol/tests/cobol/test_write_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT LF ASSIGN TO "l.txt" ORGANIZATION IS LINE SEQUENTIAL.
DATA DIVISION.
FILE SECTION.
FD LF.
01 LR PIC X(80).
PROCEDURE DIVISION.
    OPEN OUTPUT LF.
    WRITE LR.
    CLOSE LF.
    STOP RUN.

