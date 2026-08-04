*> vybe-test: cobol/math_numeric_expanded/numeric_round_trip_compute_compiles
*> origin: languages/cobol/tests/cobol/test_math_numeric_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 9.
01 B PIC 9(3) VALUE 4.
01 C PIC 9(3).
PROCEDURE DIVISION.
    COMPUTE C = A - B.
    ADD C TO B.
    STOP RUN.

