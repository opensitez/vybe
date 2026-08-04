*> vybe-test: cobol/datetime_and_encoding/use_function_when_compiled
*> origin: languages/cobol/tests/cobol/test_datetime_and_encoding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(8).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO WS-DATE.
    STOP RUN.

