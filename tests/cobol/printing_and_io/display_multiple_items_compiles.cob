*> vybe-test: cobol/printing_and_io/display_multiple_items_compiles
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10) VALUE "BOB".
01 WS-AGE PIC 9(2) VALUE 42.
PROCEDURE DIVISION.
    DISPLAY "Name: " WS-NAME " Age: " WS-AGE.
    STOP RUN.

