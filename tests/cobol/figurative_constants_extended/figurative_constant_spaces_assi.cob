*> vybe-test: cobol/figurative_constants_extended/figurative_constant_spaces_assign_blank_value
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(5) VALUE "HELLO".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE SPACES TO WS-NAME.
    DISPLAY WS-NAME.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-NAME DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "     "
        DISPLAY "FAIL: want [     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

