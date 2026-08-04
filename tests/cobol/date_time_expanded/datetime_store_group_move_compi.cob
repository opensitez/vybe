*> vybe-test: cobol/date_time_expanded/datetime_store_group_move_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(8).
01 TS.
   05 Y PIC 9(4).
   05 M PIC 9(2).
   05 DD PIC 9(2).
PROCEDURE DIVISION.
    ACCEPT D FROM DATE.
    MOVE D TO TS.
    STOP RUN.

