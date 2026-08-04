*> vybe-test: cobol/condition_names_level88_states/condition_name_runtime_recomputed_after_move
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 0 TO S.
    IF ST-A DISPLAY "A" ELSE DISPLAY "Z" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Z"
        DISPLAY "FAIL: want [Z] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

