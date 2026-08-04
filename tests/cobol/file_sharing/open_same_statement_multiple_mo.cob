*> vybe-test: cobol/file_sharing/open_same_statement_multiple_modes_compiles
*> origin: languages/cobol/tests/cobol/test_file_sharing.rs
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
    OPEN INPUT F1 OUTPUT F2.
    CLOSE F1 F2.
    STOP RUN.

