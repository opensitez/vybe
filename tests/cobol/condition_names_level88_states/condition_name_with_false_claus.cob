*> vybe-test: cobol/condition_names_level88_states/condition_name_with_false_clause_set_false_updates_storage_and_condition
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SW PIC 9 VALUE 1.
   88 ENABLED VALUE 1 WHEN SET TO FALSE IS 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET ENABLED TO FALSE.
    DISPLAY SW.
    MOVE SPACES TO WS-VYBE-L
    STRING SW DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    IF ENABLED DISPLAY "Y" ELSE DISPLAY "N" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "Y" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "N"
        DISPLAY "FAIL: want [N] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

