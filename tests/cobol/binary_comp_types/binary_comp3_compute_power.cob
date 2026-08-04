*> vybe-test: cobol/binary_comp_types/binary_comp3_compute_power
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BASE PIC 9(3) COMP-3 VALUE 4.
01 R PIC 9(5) COMP-3 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = BASE ** 3.
    STOP RUN.

