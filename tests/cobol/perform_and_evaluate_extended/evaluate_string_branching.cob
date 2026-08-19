*> vybe-test: cobol/perform_and_evaluate_extended/evaluate_string_branching
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC X VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE WS-CODE
        WHEN "A"
            DISPLAY "ALPHA"
        WHEN "B"
            DISPLAY "BETA"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ALPHA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALPHA"
        DISPLAY "FAIL: want [ALPHA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

