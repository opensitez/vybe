*> vybe-test: cobol/condition_names_level88_states/condition_name_with_multiple_values_true_after_move
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC 9 VALUE 0.
   88 OK-STATE VALUE 1 2 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 2 TO ST.
    IF OK-STATE DISPLAY "OK" ELSE DISPLAY "BAD" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "OK" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

