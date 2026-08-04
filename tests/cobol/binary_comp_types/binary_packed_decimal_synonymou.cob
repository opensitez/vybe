*> vybe-test: cobol/binary_comp_types/binary_packed_decimal_synonymous_with_comp3
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(7) PACKED-DECIMAL VALUE 0.
PROCEDURE DIVISION.
    ADD 100 TO N.
    STOP RUN.

