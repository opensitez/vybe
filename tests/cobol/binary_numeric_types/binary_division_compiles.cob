*> vybe-test: cobol/binary_numeric_types/binary_division_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) COMP VALUE 2.
01 B PIC S9(4) COMP VALUE 8.
PROCEDURE DIVISION.
    DIVIDE A INTO B.
    STOP RUN.

