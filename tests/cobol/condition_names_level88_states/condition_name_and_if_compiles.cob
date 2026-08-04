*> vybe-test: cobol/condition_names_level88_states/condition_name_and_if_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
PROCEDURE DIVISION.
    IF A = 1 AND ST-A DISPLAY "Y" END-IF.
    STOP RUN.

