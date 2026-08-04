*> vybe-test: cobol/condition_names_level88_states/condition_name_with_display_chain_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
PROCEDURE DIVISION.
    IF ST-A DISPLAY "A" "B" END-IF.
    STOP RUN.

