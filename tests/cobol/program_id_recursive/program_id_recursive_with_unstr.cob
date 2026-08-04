*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_unstring
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(10) VALUE "A,B".
01 F1 PIC X(5).
01 F2 PIC X(5).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2.
    STOP RUN.

