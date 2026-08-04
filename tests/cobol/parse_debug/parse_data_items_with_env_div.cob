*> vybe-test: cobol/parse_debug/parse_data_items_with_env_div
*> origin: languages/cobol/tests/cobol/test_parse_debug.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CALC.
ENVIRONMENT DIVISION.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MSG PIC X(20) VALUE "From COBOL".
PROCEDURE DIVISION.
    DISPLAY WS-MSG.
    STOP RUN.

