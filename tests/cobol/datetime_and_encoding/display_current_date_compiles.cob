*> vybe-test: cobol/datetime_and_encoding/display_current_date_compiles
*> origin: languages/cobol/tests/cobol/test_datetime_and_encoding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    DISPLAY CURRENT-DATE.
    STOP RUN.

