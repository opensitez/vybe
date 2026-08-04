*> vybe-test: cobol/select_assign/select_assign_to_disk_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT CUST ASSIGN TO DISK.
DATA DIVISION.
FILE SECTION.
FD CUST.
01 R PIC X(10).
PROCEDURE DIVISION.
    STOP RUN.

