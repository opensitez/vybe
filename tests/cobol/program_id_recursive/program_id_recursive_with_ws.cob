*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_ws
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO N.
    STOP RUN.

