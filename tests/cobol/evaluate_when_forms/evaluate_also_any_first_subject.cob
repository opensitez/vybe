*> vybe-test: cobol/evaluate_when_forms/evaluate_also_any_first_subject_matches_all
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 7.
01 B PIC 9 VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN ANY ALSO ANY
            DISPLAY "CATCH-ALL"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "CATCH-ALL" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CATCH-ALL"
        DISPLAY "FAIL: want [CATCH-ALL] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

