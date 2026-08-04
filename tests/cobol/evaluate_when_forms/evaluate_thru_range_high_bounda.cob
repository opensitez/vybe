*> vybe-test: cobol/evaluate_when_forms/evaluate_thru_range_high_boundary
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(2) VALUE 20.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 10 THRU 20
            DISPLAY "IN"
        WHEN OTHER
            DISPLAY "OUT"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "IN" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "IN"
        DISPLAY "FAIL: want [IN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

