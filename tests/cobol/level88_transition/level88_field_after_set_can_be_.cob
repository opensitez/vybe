*> vybe-test: cobol/level88_transition/level88_field_after_set_can_be_tested_multiple_times
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X VALUE "N".
    88 YES VALUE "Y".
    88 NO-FLAG VALUE "N".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NO-FLAG
        SET YES TO TRUE
    END-IF.
    IF YES
        DISPLAY "YES"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "YES"
        DISPLAY "FAIL: want [YES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

