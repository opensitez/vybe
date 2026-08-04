*> vybe-test: cobol/procedure_division_expanded/set_false_condition_name_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1).
   88 WS-OFF VALUE 0.
PROCEDURE DIVISION.
    SET WS-OFF TO FALSE.
    STOP RUN.

