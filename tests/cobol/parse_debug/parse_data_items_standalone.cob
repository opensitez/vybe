*> vybe-test: cobol/parse_debug/parse_data_items_standalone
*> origin: languages/cobol/tests/cobol/test_parse_debug.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC X(10) VALUE "Hello".
PROCEDURE DIVISION.
    STOP RUN.

