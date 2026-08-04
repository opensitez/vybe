*> vybe-test: cobol/floating_point/floating_multiply_compiles
*> origin: languages/cobol/tests/cobol/test_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FP8.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE COMP-1.
01 B USAGE COMP-1.
PROCEDURE DIVISION.
    MULTIPLY A BY B.
    STOP RUN.

