*> vybe-test: cobol/math_numeric_expanded/add_giving_compiles
*> origin: languages/cobol/tests/cobol/test_math_numeric_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 1.
01 B PIC 9(3) VALUE 2.
01 C PIC 9(3).
PROCEDURE DIVISION.
    ADD A B GIVING C.
    STOP RUN.

