*> vybe-test: cobol/international_text_support/national_text_move_and_display_compiles
*> origin: languages/cobol/tests/cobol/test_international_text_support.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(20) USAGE NATIONAL VALUE "Unicode".
01 WS-DST PIC X(20) USAGE NATIONAL.
PROCEDURE DIVISION.
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
    STOP RUN.

