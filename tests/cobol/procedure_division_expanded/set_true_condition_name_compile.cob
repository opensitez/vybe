*> vybe-test: cobol/procedure_division_expanded/set_true_condition_name_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1).
   88 WS-ON VALUE 1.
PROCEDURE DIVISION.
    SET WS-ON TO TRUE.
    STOP RUN.

