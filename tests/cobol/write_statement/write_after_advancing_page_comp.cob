*> vybe-test: cobol/write_statement/write_after_advancing_page_compiles
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
    WRITE R AFTER ADVANCING PAGE.
    CLOSE F.
    STOP RUN.

