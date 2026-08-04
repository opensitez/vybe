*> vybe-test: cobol/encoding_international_text/national_move_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(20) USAGE NATIONAL VALUE "TXT".
01 B PIC X(20) USAGE NATIONAL.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

