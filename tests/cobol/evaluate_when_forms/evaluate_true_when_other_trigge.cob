*> vybe-test: cobol/evaluate_when_forms/evaluate_true_when_other_triggered
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN N > 10
            DISPLAY "BIG"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "BIG" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OTHER"
        DISPLAY "FAIL: want [OTHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

