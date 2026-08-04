*> vybe-test: cobol/math_numeric_expanded/divide_basic_compiles
*> origin: languages/cobol/tests/cobol/test_math_numeric_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 8.
01 B PIC 9(3) VALUE 2.
PROCEDURE DIVISION.
    DIVIDE A BY B.
    STOP RUN.

