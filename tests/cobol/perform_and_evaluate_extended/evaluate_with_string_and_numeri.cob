*> vybe-test: cobol/perform_and_evaluate_extended/evaluate_with_string_and_numeric_alternatives
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC X VALUE "C".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE WS-CODE
        WHEN "A"
            DISPLAY "ALPHA"
        WHEN "C"
            DISPLAY "CHARLIE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ALPHA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CHARLIE"
        DISPLAY "FAIL: want [CHARLIE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

