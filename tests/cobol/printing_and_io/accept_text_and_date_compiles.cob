*> vybe-test: cobol/printing_and_io/accept_text_and_date_compiles
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20).
01 WS-DATE PIC X(8).
PROCEDURE DIVISION.
    ACCEPT WS-NAME.
    ACCEPT WS-DATE FROM DATE.
    STOP RUN.

