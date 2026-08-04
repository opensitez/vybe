*> vybe-test: cobol/rewrite_statement/rewrite_sequential_io_compiles
*> origin: languages/cobol/tests/cobol/test_rewrite_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "s.dat" ORGANIZATION IS SEQUENTIAL.
DATA DIVISION.
FILE SECTION.
FD F.
01 REC PIC X(20).
PROCEDURE DIVISION.
    OPEN I-O F.
    REWRITE REC.
    CLOSE F.
    STOP RUN.

