*> vybe-test: cobol/conditions_extended/if_spaces_runtime
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC X(4) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF WS-X IS SPACES
        DISPLAY "SPACES"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "SPACES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SPACES"
        DISPLAY "FAIL: want [SPACES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

