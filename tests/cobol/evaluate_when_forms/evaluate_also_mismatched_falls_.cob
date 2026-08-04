*> vybe-test: cobol/evaluate_when_forms/evaluate_also_mismatched_falls_through_to_other
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN 1 ALSO 2
            DISPLAY "1-2"
        WHEN OTHER ALSO OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "1-2" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OTHER"
        DISPLAY "FAIL: want [OTHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

