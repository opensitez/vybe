*> vybe-test: cobol/program_id_recursive/program_id_recursive_compute_chain
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 5.
01 Y PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE Y = X ** 2 + 2 * X + 1.
    STOP RUN.

