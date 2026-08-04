*> vybe-test: cobol/date_time_expanded/current_date_to_group_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TS.
   05 Y PIC 9(4).
   05 M PIC 9(2).
   05 D PIC 9(2).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE(1:8) TO TS.
    STOP RUN.

