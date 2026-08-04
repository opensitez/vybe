*> vybe-test: cobol/evaluate_when_forms/evaluate_grade_classification
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SCORE PIC 9(3) VALUE 85.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE SCORE
        WHEN 90 THRU 100
            DISPLAY "A"
        WHEN 80 THRU 89
            DISPLAY "B"
        WHEN 70 THRU 79
            DISPLAY "C"
        WHEN OTHER
            DISPLAY "F"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

