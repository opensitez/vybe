*> vybe-test: cobol/encoding_international_text/json_unicode_data_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 J PIC X(100).
01 R PIC X(20) USAGE NATIONAL.
PROCEDURE DIVISION.
    JSON PARSE J INTO R.
    STOP RUN.

