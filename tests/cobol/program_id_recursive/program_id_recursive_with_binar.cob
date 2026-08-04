*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_binary_fields
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 COUNT PIC 9(9) COMP VALUE 0.
01 TOTAL PIC 9(12) COMP-3 VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO COUNT.
    ADD 100 TO TOTAL.
    STOP RUN.

