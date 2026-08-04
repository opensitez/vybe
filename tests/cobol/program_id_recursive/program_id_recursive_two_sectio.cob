*> vybe-test: cobol/program_id_recursive/program_id_recursive_two_sections
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM SEC-A.
    PERFORM SEC-B.
    STOP RUN.
SEC-A SECTION.
    ADD 1 TO X.
SEC-B SECTION.
    ADD 2 TO X.
    STOP RUN.

