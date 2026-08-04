*> vybe-test: cobol/binary_comp_types/binary_comp2_double_compiles
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D COMP-2 VALUE 3.14159.
PROCEDURE DIVISION.
    COMPUTE D = D + 1.
    STOP RUN.

