*> vybe-test: cobol/select_assign/select_assign_with_file_status_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT CUST ASSIGN TO "cust.dat" FILE STATUS IS FS.
DATA DIVISION.
FILE SECTION.
FD CUST.
01 R PIC X(10).
WORKING-STORAGE SECTION.
01 FS PIC XX.
PROCEDURE DIVISION.
    STOP RUN.

