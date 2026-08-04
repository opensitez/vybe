*> vybe-test: cobol/date_time_expanded/date_intrinsic_length_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(8).
01 L PIC 9(3).
PROCEDURE DIVISION.
    ACCEPT D FROM DATE.
    MOVE FUNCTION LENGTH(D) TO L.
    STOP RUN.

