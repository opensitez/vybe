*> vybe-test: cobol/binary_comp_types/binary_comp5_basic_value
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(9) COMP-5 VALUE 0.
PROCEDURE DIVISION.
    ADD 42 TO N.
    STOP RUN.

