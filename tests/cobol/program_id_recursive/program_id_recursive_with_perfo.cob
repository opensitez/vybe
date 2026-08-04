*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_perform_para
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    PERFORM CALC.
    STOP RUN.
CALC.
    ADD 42 TO R.
    STOP RUN.

