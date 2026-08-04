*> vybe-test: cobol/date_time_expanded/date_store_and_move_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D1 PIC X(8).
01 D2 PIC X(8).
PROCEDURE DIVISION.
    ACCEPT D1 FROM DATE.
    MOVE D1 TO D2.
    STOP RUN.

