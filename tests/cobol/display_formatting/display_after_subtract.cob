*> vybe-test: cobol/display_formatting/display_after_subtract
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 200.
01 B PIC 9(3) VALUE 75.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "125"
        DISPLAY "FAIL: want [125] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

