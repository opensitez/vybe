*> vybe-test: cobol/level88_transition/level88_evaluate_true_multiple_when
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 CODE PIC 9 VALUE 2.
    88 CODE-ONE VALUE 1.
    88 CODE-TWO VALUE 2.
    88 CODE-THREE VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN CODE-ONE
            DISPLAY "ONE"
        WHEN CODE-TWO
            DISPLAY "TWO"
        WHEN CODE-THREE
            DISPLAY "THREE"
        WHEN OTHER
            DISPLAY "MANY"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "TWO"
        DISPLAY "FAIL: want [TWO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

