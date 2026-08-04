*> vybe-test: cobol/condition_names_level88_states/condition_name_multi_values_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
   88 ST-B VALUE 2.
PROCEDURE DIVISION.
    IF ST-A DISPLAY "A" END-IF.
    STOP RUN.

