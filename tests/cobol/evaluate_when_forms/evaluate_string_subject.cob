*> vybe-test: cobol/evaluate_when_forms/evaluate_string_subject
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(3) VALUE "YES".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE S
        WHEN "YES"
            DISPLAY "AFFIRMATIVE"
        WHEN "NO"
            DISPLAY "NEGATIVE"
        WHEN OTHER
            DISPLAY "UNKNOWN"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "AFFIRMATIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AFFIRMATIVE"
        DISPLAY "FAIL: want [AFFIRMATIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

