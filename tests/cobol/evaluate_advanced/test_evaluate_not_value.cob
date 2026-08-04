*> vybe-test: cobol/evaluate_advanced/test_evaluate_not_value
*> origin: languages/cobol/tests/cobol/test_evaluate_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VAL PIC 9 VALUE 4.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE WS-VAL
        WHEN NOT 3
            DISPLAY "NOT-THREE"
        WHEN OTHER
            DISPLAY "THREE"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "NOT-THREE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NOT-THREE"
        DISPLAY "FAIL: want [NOT-THREE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

