*> vybe-test: cobol/math_builtins/compute_multiplication_compiles
*> origin: languages/cobol/tests/cobol/test_math_builtins.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 6.
01 WS-B PIC 9(3) VALUE 7.
01 WS-C PIC 9(3).
PROCEDURE DIVISION.
    COMPUTE WS-C = WS-A * WS-B.
    STOP RUN.

