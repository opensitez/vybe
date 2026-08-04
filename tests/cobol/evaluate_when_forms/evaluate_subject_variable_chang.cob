*> vybe-test: cobol/evaluate_when_forms/evaluate_subject_variable_changed_mid_program
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 STATUS PIC X VALUE "A".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "B" TO STATUS.
    EVALUATE STATUS
        WHEN "A"
            DISPLAY "A"
        WHEN "B"
            DISPLAY "B"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

