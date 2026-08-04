*> vybe-test: cobol/display_advanced/test_display_signed_negative
*> origin: languages/cobol/tests/cobol/test_display_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
01 WS-NUM PIC S9(3) VALUE -42.
    STOP RUN.

