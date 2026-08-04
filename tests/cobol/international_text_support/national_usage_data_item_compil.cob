*> vybe-test: cobol/international_text_support/national_usage_data_item_compiles
*> origin: languages/cobol/tests/cobol/test_international_text_support.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(50) USAGE NATIONAL.
PROCEDURE DIVISION.
    DISPLAY WS-TEXT.
    STOP RUN.

