*> vybe-test: cobol/condition_names_level88_states/condition_name_loop_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
PROCEDURE DIVISION.
    PERFORM UNTIL N >= 2
        ADD 1 TO N
        IF ST-A DISPLAY "A" END-IF
    END-PERFORM.
    STOP RUN.

