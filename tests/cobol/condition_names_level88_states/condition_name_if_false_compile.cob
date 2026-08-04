*> vybe-test: cobol/condition_names_level88_states/condition_name_if_false_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 0.
   88 ONN VALUE 1.
PROCEDURE DIVISION.
    IF NOT ONN DISPLAY "N" END-IF.
    STOP RUN.

