*> vybe-test: cobol/select_assign/select_assign_line_sequential_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT TXT ASSIGN TO "cust.txt" ORGANIZATION IS LINE SEQUENTIAL.
DATA DIVISION.
FILE SECTION.
FD TXT.
01 R PIC X(80).
PROCEDURE DIVISION.
    STOP RUN.

