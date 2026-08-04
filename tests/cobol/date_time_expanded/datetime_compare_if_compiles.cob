*> vybe-test: cobol/date_time_expanded/datetime_compare_if_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC X(8).
PROCEDURE DIVISION.
    ACCEPT T FROM TIME.
    IF T > "12000000"
        DISPLAY "PM"
    END-IF.
    STOP RUN.

