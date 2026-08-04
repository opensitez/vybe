*> vybe-test: cobol/level88_transition/level88_multiple_values
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 GRADE PIC X VALUE "B".
    88 PASSING VALUE "A" "B" "C".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF PASSING
        DISPLAY "PASS"
    ELSE
        DISPLAY "FAIL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "PASS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "PASS"
        DISPLAY "FAIL: want [PASS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

