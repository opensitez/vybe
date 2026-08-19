*> vybe-test: cobol/evaluate_advanced/test_evaluate_grading_ranges
*> origin: languages/cobol/tests/cobol/test_evaluate_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SCORE PIC 9(3) VALUE 85.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE TRUE
        WHEN WS-SCORE >= 90
            DISPLAY "A"
        WHEN WS-SCORE >= 80
            DISPLAY "B"
        WHEN WS-SCORE >= 70
            DISPLAY "C"
        WHEN OTHER
            DISPLAY "F"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

