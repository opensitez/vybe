*> vybe-test: cobol/evaluate_advanced/test_evaluate_expr_subject
*> origin: languages/cobol/tests/cobol/test_evaluate_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 1.
01 WS-B PIC 9 VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE WS-A + WS-B
        WHEN 3
            DISPLAY "THREE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "THREE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "THREE"
        DISPLAY "FAIL: want [THREE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

