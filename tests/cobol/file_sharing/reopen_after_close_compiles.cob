*> vybe-test: cobol/file_sharing/reopen_after_close_compiles
*> origin: languages/cobol/tests/cobol/test_file_sharing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F1 ASSIGN TO "a.dat".
DATA DIVISION.
FILE SECTION.
FD F1.
01 R1 PIC X(10).
PROCEDURE DIVISION.
    OPEN INPUT F1.
    CLOSE F1.
    OPEN INPUT F1.
    CLOSE F1.
    STOP RUN.

