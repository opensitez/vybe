*> vybe-test: cobol/level88_transition/level88_nested_if_with_multiple_flags
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A-FLAG PIC X VALUE "Y".
    88 A-ON VALUE "Y".
01 B-FLAG PIC X VALUE "N".
    88 B-ON VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A-ON
        IF B-ON
            DISPLAY "BOTH"
        ELSE
            DISPLAY "ONLY A"
        END-IF
    ELSE
        DISPLAY "NOT A"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BOTH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONLY A"
        DISPLAY "FAIL: want [ONLY A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

