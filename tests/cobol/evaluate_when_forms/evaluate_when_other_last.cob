*> vybe-test: cobol/evaluate_when_forms/evaluate_when_other_last
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
        WHEN OTHER
            DISPLAY "MANY"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE"
        DISPLAY "FAIL: want [ONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

