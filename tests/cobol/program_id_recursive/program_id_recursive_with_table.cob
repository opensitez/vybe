*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_table
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(2) OCCURS 5 TIMES.
PROCEDURE DIVISION.
    MOVE 42 TO E(3).
    STOP RUN.

