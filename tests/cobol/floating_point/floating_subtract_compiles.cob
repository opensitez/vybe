*> vybe-test: cobol/floating_point/floating_subtract_compiles
*> origin: languages/cobol/tests/cobol/test_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FP7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE COMP-2.
01 B USAGE COMP-2.
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    STOP RUN.

