*> vybe-test: cobol/evaluate_when_forms/evaluate_first_match_wins
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 1 THRU 10
            DISPLAY "FIRST"
        WHEN 5
            DISPLAY "SECOND"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "FIRST" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FIRST"
        DISPLAY "FAIL: want [FIRST] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

