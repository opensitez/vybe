*> vybe-test: cobol/compute_rounded/compute_rounded_negative_result
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC S9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R ROUNDED = -3.7.
    STOP RUN.

