*> vybe-test: cobol/condition_names_level88_states/condition_name_runtime_boolean_composition_with_and
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC 9 VALUE 1.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF FLAG = 1 AND ST-A DISPLAY "BOTH" ELSE DISPLAY "MISS" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BOTH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BOTH"
        DISPLAY "FAIL: want [BOTH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

