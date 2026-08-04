*> vybe-test: cobol/compute_rounded/compute_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R ROUNDED = 7 / 3.
    STOP RUN.

