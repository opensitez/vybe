*> vybe-test: cobol/condition_names_level88_states/condition_name_set_true_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9.
   88 ONN VALUE 1.
PROCEDURE DIVISION.
    SET ONN TO TRUE.
    STOP RUN.

