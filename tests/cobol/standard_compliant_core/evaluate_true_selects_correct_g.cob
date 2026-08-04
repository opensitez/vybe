*> vybe-test: cobol/standard_compliant_core/evaluate_true_selects_correct_grade_band
*> origin: languages/cobol/tests/cobol/test_standard_compliant_core.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SCORE PIC 9(3) VALUE 82.
01 WS-GRADE PIC X VALUE "?".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE TRUE
        WHEN WS-SCORE >= 90
            MOVE "A" TO WS-GRADE
        WHEN WS-SCORE >= 80
            MOVE "B" TO WS-GRADE
        WHEN OTHER
            MOVE "C" TO WS-GRADE
    END-EVALUATE.
    DISPLAY WS-GRADE.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-GRADE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

