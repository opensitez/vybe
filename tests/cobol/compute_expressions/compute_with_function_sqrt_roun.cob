*> vybe-test: cobol/compute_expressions/compute_with_function_sqrt_rounded
*> origin: languages/cobol/tests/cobol/test_compute_expressions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(5)V99 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION SQRT(144).
    STOP RUN.

