*> vybe-test: cobol/condition_names_level88_states/condition_name_move_and_set_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 0.
   88 ST-A VALUE 1.
PROCEDURE DIVISION.
    MOVE 1 TO S.
    IF ST-A DISPLAY "A" END-IF.
    STOP RUN.

