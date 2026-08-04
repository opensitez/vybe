*> vybe-test: cobol/evaluate_advanced/test_evaluate_false
*> origin: languages/cobol/tests/cobol/test_evaluate_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE FALSE
        WHEN WS-A = 5
            DISPLAY "NOT-FIVE"
        WHEN OTHER
            DISPLAY "FIVE"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "NOT-FIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FIVE"
        DISPLAY "FAIL: want [FIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

