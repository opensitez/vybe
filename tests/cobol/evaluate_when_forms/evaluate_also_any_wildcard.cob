*> vybe-test: cobol/evaluate_when_forms/evaluate_also_any_wildcard
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN ANY ALSO 3
            DISPLAY "B IS 3"
        WHEN OTHER ALSO OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "B IS 3" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B IS 3"
        DISPLAY "FAIL: want [B IS 3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

