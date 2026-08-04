*> vybe-test: cobol/display_formatting/display_upon_console_compiles
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    DISPLAY "OUTPUT" UPON CONSOLE.
    STOP RUN.

