*> vybe-test: cobol/display_advanced/test_display_special_destinations
*> origin: languages/cobol/tests/cobol/test_display_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    DISPLAY "HELLO" UPON CONSOLE.
    DISPLAY "ERROR" UPON SYSERR.
    STOP RUN.

