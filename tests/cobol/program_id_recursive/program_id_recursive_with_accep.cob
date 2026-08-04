*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_accept_from_date
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TODAY PIC 9(6).
PROCEDURE DIVISION.
    ACCEPT TODAY FROM DATE.
    STOP RUN.

