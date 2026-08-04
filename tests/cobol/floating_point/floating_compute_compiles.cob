*> vybe-test: cobol/floating_point/floating_compute_compiles
*> origin: languages/cobol/tests/cobol/test_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FP3.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE COMP-1.
01 B USAGE COMP-2.
01 C USAGE COMP-2.
PROCEDURE DIVISION.
    COMPUTE C = A + B.
    STOP RUN.

