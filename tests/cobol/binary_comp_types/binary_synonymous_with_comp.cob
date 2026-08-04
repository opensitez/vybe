*> vybe-test: cobol/binary_comp_types/binary_synonymous_with_comp
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(8) BINARY VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO N.
    STOP RUN.

