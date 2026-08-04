*> vybe-test: cobol/level88_transition/level88_set_to_true
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "N".
    88 FLAG-ON VALUE "Y".
    88 FLAG-OFF VALUE "N".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET FLAG-ON TO TRUE.
    IF FLAG-ON
        DISPLAY "ON"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ON" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ON"
        DISPLAY "FAIL: want [ON] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

