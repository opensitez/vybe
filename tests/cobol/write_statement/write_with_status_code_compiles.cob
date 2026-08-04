*> vybe-test: cobol/write_statement/write_with_status_code_compiles
*> origin: languages/cobol/tests/cobol/test_write_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat" FILE STATUS IS FS.
DATA DIVISION.
FILE SECTION.
FD F.
01 R PIC X(20).
WORKING-STORAGE SECTION.
01 FS PIC XX.
PROCEDURE DIVISION.
    OPEN OUTPUT F.
    WRITE R.
    CLOSE F.
    STOP RUN.

