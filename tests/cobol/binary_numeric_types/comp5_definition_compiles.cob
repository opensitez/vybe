*> vybe-test: cobol/binary_numeric_types/comp5_definition_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN3.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 B1 PIC S9(9) COMP-5.
PROCEDURE DIVISION.
    STOP RUN.

