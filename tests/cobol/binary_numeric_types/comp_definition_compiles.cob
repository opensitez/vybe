*> vybe-test: cobol/binary_numeric_types/comp_definition_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 B1 PIC S9(4) COMP.
PROCEDURE DIVISION.
    STOP RUN.

