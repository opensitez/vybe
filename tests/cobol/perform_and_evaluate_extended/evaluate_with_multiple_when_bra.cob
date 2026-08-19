*> vybe-test: cobol/perform_and_evaluate_extended/evaluate_with_multiple_when_branches
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9 VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE WS-X
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE"
        DISPLAY "FAIL: want [ONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

