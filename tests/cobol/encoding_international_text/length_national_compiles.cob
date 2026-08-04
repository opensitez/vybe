*> vybe-test: cobol/encoding_international_text/length_national_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20) VALUE "A".
01 L PIC 9(3).
PROCEDURE DIVISION.
    MOVE FUNCTION LENGTH(N) TO L.
    STOP RUN.

