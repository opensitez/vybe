*> vybe-test: cobol/evaluate_when_forms/evaluate_true_compound_when_condition
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN A = 5 AND B = 5
            DISPLAY "BOTH FIVE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "BOTH FIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BOTH FIVE"
        DISPLAY "FAIL: want [BOTH FIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

