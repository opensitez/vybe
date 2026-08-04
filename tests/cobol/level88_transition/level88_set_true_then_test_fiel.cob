*> vybe-test: cobol/level88_transition/level88_set_true_then_test_field_value
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "N".
    88 YES-FLAG VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET YES-FLAG TO TRUE.
    DISPLAY FLAG.
    MOVE SPACES TO WS-VYBE-L
    STRING FLAG DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

