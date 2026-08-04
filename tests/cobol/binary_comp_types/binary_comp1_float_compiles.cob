*> vybe-test: cobol/binary_comp_types/binary_comp1_float_compiles
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F COMP-1 VALUE 3.14.
PROCEDURE DIVISION.
    COMPUTE F = F * 2.
    STOP RUN.

