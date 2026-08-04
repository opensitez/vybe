*> vybe-test: cobol/date_time_expanded/time_store_and_move_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T1 PIC X(8).
01 T2 PIC X(8).
PROCEDURE DIVISION.
    ACCEPT T1 FROM TIME.
    MOVE T1 TO T2.
    STOP RUN.

