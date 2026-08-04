*> vybe-test: cobol/printing_and_io/display_multiple_items_runtime_formats_output
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(5) VALUE "ALICE".
01 WS-AGE PIC 9(2) VALUE 31.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY "Name=" WS-NAME " Age=" WS-AGE.
    MOVE SPACES TO WS-VYBE-L
    STRING "Name=" DELIMITED SIZE WS-NAME DELIMITED SIZE " Age=" DELIMITED SIZE WS-AGE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Name=ALICE Age=31"
        DISPLAY "FAIL: want [Name=ALICE Age=31] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

