*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_string_op
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "HELLO".
01 R PIC X(15) VALUE SPACES.
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE INTO R.
    STOP RUN.

