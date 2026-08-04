*> vybe-test: cobol/encoding_international_text/national_to_display_move_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20) USAGE NATIONAL VALUE "HELLO".
01 D PIC X(20).
PROCEDURE DIVISION.
    MOVE N TO D.
    DISPLAY D.
    STOP RUN.

