*> vybe-test: cobol/encoding_international_text/national_usage_decl_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(30) USAGE NATIONAL.
PROCEDURE DIVISION.
    DISPLAY N.
    STOP RUN.

