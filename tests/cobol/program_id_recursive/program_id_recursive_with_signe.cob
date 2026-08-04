*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_signed_field
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(5) VALUE -100.
PROCEDURE DIVISION.
    ADD 200 TO N.
    STOP RUN.

