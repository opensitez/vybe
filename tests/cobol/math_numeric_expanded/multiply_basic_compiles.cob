*> vybe-test: cobol/math_numeric_expanded/multiply_basic_compiles
*> origin: languages/cobol/tests/cobol/test_math_numeric_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 3.
01 B PIC 9(3) VALUE 2.
PROCEDURE DIVISION.
    MULTIPLY A BY B.
    STOP RUN.

