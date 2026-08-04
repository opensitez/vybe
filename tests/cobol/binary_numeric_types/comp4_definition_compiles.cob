*> vybe-test: cobol/binary_numeric_types/comp4_definition_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN2.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 B1 PIC S9(4) COMP-4.
PROCEDURE DIVISION.
    STOP RUN.

