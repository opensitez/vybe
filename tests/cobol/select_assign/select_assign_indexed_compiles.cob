*> vybe-test: cobol/select_assign/select_assign_indexed_compiles
*> origin: languages/cobol/tests/cobol/test_select_assign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT IDX ASSIGN TO "i.dat" ORGANIZATION IS INDEXED RECORD KEY IS K.
DATA DIVISION.
FILE SECTION.
FD IDX.
01 REC.
   05 K PIC 9(5).
PROCEDURE DIVISION.
    STOP RUN.

