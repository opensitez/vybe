*> vybe-test: cobol/printing_and_io/display_literal_and_variable_compiles
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10) VALUE "ALICE".
PROCEDURE DIVISION.
    DISPLAY "Hello".
    DISPLAY WS-NAME.
    STOP RUN.

