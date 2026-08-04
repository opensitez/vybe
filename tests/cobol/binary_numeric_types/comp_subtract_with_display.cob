*> vybe-test: cobol/binary_numeric_types/comp_subtract_with_display
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN12.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A PIC S9(4) COMP VALUE 20.
01 B PIC S9(4) COMP VALUE 7.
PROCEDURE DIVISION.
    SUBTRACT B FROM A
    DISPLAY A
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "13"
        DISPLAY "FAIL: want [13] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

