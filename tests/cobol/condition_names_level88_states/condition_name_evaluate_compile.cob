*> vybe-test: cobol/condition_names_level88_states/condition_name_evaluate_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE S WHEN 1 DISPLAY "A" WHEN 2 DISPLAY "B" WHEN OTHER DISPLAY "X" END-EVALUATE.
    STOP RUN.

