*> vybe-test: cobol/signed_arithmetic/signed_s9_comp_compiles
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(8) COMP VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO N.
    STOP RUN.

