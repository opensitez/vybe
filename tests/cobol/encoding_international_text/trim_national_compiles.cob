*> vybe-test: cobol/encoding_international_text/trim_national_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20) USAGE NATIONAL.
01 O PIC X(20) USAGE NATIONAL.
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(N) TO O.
    STOP RUN.

