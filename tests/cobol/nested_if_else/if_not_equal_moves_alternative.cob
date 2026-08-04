*> vybe-test: cobol/nested_if_else/if_not_equal_moves_alternative
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "HELLO".
01 R PIC X(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF S NOT = "HELLO"
        MOVE "WRONG" TO R
    ELSE
        MOVE "RIGHT" TO R
    END-IF.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "RIGHT"
        DISPLAY "FAIL: want [RIGHT] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

