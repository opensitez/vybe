*> vybe-test: cobol/condition_names/condition_name_recomputed_after_move
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC X VALUE "A".
   88 IS-VALID VALUE "A".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE "B" TO WS-CODE.
    IF IS-VALID
        DISPLAY "VALID"
    ELSE
        DISPLAY "INVALID"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "VALID" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "VALID"
        DISPLAY "FAIL: want [VALID] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

