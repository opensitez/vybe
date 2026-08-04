*> vybe-test: cobol/condition_names_level88_states/condition_name_runtime_false_branch_prints_no
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 2.
   88 ST-A VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF ST-A DISPLAY "YES" ELSE DISPLAY "NO" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NO"
        DISPLAY "FAIL: want [NO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

