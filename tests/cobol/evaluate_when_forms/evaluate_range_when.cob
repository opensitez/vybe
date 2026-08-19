*> vybe-test: cobol/evaluate_when_forms/evaluate_range_when
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(2) VALUE 45.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 1 THRU 25
            DISPLAY "LOW"
        WHEN 26 THRU 75
            DISPLAY "MID"
        WHEN 76 THRU 99
            DISPLAY "HIGH"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "LOW" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LOW"
        DISPLAY "FAIL: want [LOW] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

