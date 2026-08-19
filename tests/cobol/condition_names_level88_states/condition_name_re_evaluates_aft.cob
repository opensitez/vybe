*> vybe-test: cobol/condition_names_level88_states/condition_name_re_evaluates_after_arithmetic_change
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC 9 VALUE 1.
   88 READY VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD 1 TO ST.
    IF READY DISPLAY "READY" ELSE DISPLAY "NOT-READY" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "READY" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "READY"
        DISPLAY "FAIL: want [READY] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

