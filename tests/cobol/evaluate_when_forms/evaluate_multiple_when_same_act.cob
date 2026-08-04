*> vybe-test: cobol/evaluate_when_forms/evaluate_multiple_when_same_action
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC X VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE C
        WHEN "A"
        WHEN "B"
            DISPLAY "A OR B"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A OR B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A OR B"
        DISPLAY "FAIL: want [A OR B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

