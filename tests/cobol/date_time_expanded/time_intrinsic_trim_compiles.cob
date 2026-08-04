*> vybe-test: cobol/date_time_expanded/time_intrinsic_trim_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC X(8).
01 O PIC X(8).
PROCEDURE DIVISION.
    ACCEPT T FROM TIME.
    MOVE FUNCTION TRIM(T) TO O.
    STOP RUN.

