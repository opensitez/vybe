*> vybe-test: cobol/date_time_expanded/date_time_concat_display_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(8).
01 T PIC X(8).
01 DT PIC X(20).
PROCEDURE DIVISION.
    ACCEPT D FROM DATE.
    ACCEPT T FROM TIME.
    STRING D DELIMITED BY SIZE T DELIMITED BY SIZE INTO DT.
    DISPLAY DT.
    STOP RUN.

