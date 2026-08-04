*> vybe-test: cobol/encoding_international_text/lower_national_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20) VALUE "ABC".
01 O PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION LOWER-CASE(N) TO O.
    STOP RUN.

