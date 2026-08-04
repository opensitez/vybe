*> vybe-test: cobol/compute_rounded/add_rounded_default
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V9 VALUE 10.5.
01 R PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    ADD 1.5 TO A ROUNDED.
    STOP RUN.

