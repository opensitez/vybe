*> vybe-test: cobol/evaluate_when_forms/evaluate_also_two_subjects
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN 1 ALSO 2
            DISPLAY "ONE-TWO"
        WHEN OTHER ALSO OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE-TWO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE-TWO"
        DISPLAY "FAIL: want [ONE-TWO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

