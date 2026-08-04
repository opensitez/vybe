*> vybe-test: cobol/evaluate_when_forms/evaluate_empty_string_matches_spaces
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE S
        WHEN SPACES
            DISPLAY "BLANK"
        WHEN OTHER
            DISPLAY "NON-BLANK"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "BLANK" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BLANK"
        DISPLAY "FAIL: want [BLANK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

