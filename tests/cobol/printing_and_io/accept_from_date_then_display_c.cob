*> vybe-test: cobol/printing_and_io/accept_from_date_then_display_compiles
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(8).
PROCEDURE DIVISION.
    ACCEPT WS-DATE FROM DATE YYYYMMDD.
    DISPLAY WS-DATE.
    STOP RUN.

