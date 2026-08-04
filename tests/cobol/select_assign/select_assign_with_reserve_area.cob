*> vybe-test: cobol/select_assign/select_assign_with_reserve_areas_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT CUST ASSIGN TO "cust.dat" RESERVE 1 AREAS.
DATA DIVISION.
FILE SECTION.
FD CUST.
01 R PIC X(10).
PROCEDURE DIVISION.
    STOP RUN.

