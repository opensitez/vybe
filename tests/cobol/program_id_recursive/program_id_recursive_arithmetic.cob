*> vybe-test: cobol/program_id_recursive/program_id_recursive_arithmetic_sequence
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(4) VALUE 10.
01 B PIC 9(4) VALUE 20.
01 R PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    ADD A B GIVING R.
    SUBTRACT A FROM R.
    MULTIPLY 2 BY R.
    DIVIDE 4 INTO R.
    STOP RUN.

