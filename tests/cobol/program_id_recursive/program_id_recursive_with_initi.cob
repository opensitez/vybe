*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_initialize
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "HELLO".
01 N PIC 9(5) VALUE 99999.
PROCEDURE DIVISION.
    INITIALIZE S N.
    STOP RUN.

