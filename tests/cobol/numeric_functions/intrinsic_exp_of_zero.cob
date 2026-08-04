*> vybe-test: cobol/numeric_functions/intrinsic_exp_of_zero
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(5)V9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION EXP(0).
    STOP RUN.

