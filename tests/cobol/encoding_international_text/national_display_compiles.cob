*> vybe-test: cobol/encoding_international_text/national_display_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20) USAGE NATIONAL VALUE "HELLO".
PROCEDURE DIVISION.
    DISPLAY N.
    STOP RUN.

