*> vybe-test: cobol/display_formatting/display_with_no_advancing_compiles
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    DISPLAY "PROMPT" WITH NO ADVANCING.
    STOP RUN.

