*> vybe-test: cobol/evaluate_when_forms/evaluate_true_not_condition_branch
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN NOT N > 0
            DISPLAY "ZERO OR NEG"
        WHEN OTHER
            DISPLAY "POS"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ZERO OR NEG" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZERO OR NEG"
        DISPLAY "FAIL: want [ZERO OR NEG] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

