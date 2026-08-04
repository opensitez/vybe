*> vybe-test: cobol/display_formatting/display_numeric_after_add
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 100.
01 B PIC 9(3) VALUE 55.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD B TO A.
    DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "155"
        DISPLAY "FAIL: want [155] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

