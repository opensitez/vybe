*> vybe-test: cobol/level88_transition/level88_in_evaluate_when
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 STATUS PIC X VALUE "A".
    88 IS-OPEN VALUE "A".
    88 IS-CLOSED VALUE "C".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN IS-OPEN
            DISPLAY "OPEN"
        WHEN IS-CLOSED
            DISPLAY "CLOSED"
        WHEN OTHER
            DISPLAY "UNKNOWN"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "OPEN" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OPEN"
        DISPLAY "FAIL: want [OPEN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

