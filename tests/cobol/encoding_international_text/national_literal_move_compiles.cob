*> vybe-test: cobol/encoding_international_text/national_literal_move_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC X(20) USAGE NATIONAL VALUE "HELLO".
01 N2 PIC X(20) USAGE NATIONAL.
PROCEDURE DIVISION.
    MOVE N1 TO N2.
    DISPLAY N2.
    STOP RUN.

