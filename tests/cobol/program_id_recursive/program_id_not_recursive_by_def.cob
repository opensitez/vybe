*> vybe-test: cobol/program_id_recursive/program_id_not_recursive_by_default_compiles
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO N.
    STOP RUN.

