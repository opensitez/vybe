*> vybe-test: cobol/select_assign/select_optional_assign_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT OPTIONAL CUST ASSIGN TO "cust.dat".
DATA DIVISION.
FILE SECTION.
FD CUST.
01 R PIC X(10).
PROCEDURE DIVISION.
    STOP RUN.

