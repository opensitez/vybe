*> vybe-test: cobol/perform_and_evaluate_extended/evaluate_true_with_range_conditions
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 99 VALUE 22.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE TRUE
        WHEN WS-AGE < 13
            DISPLAY "CHILD"
        WHEN WS-AGE < 20
            DISPLAY "TEEN"
        WHEN OTHER
            DISPLAY "ADULT"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "CHILD" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CHILD"
        DISPLAY "FAIL: want [CHILD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

