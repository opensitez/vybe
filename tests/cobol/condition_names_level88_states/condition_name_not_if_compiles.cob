*> vybe-test: cobol/condition_names_level88_states/condition_name_not_if_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 0.
   88 ST-A VALUE 1.
PROCEDURE DIVISION.
    IF NOT ST-A DISPLAY "N" END-IF.
    STOP RUN.

