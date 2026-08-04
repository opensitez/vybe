*> vybe-test: cobol/display_advanced/test_display_no_advancing
*> origin: languages/cobol/tests/cobol/test_display_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    DISPLAY "HELLO " WITH NO ADVANCING.
    DISPLAY "WORLD".
    STOP RUN.

