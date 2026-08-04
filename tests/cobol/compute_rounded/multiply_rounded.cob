*> vybe-test: cobol/compute_rounded/multiply_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V99 VALUE 3.33.
01 R PIC 9(4)V9 VALUE 0.
PROCEDURE DIVISION.
    MULTIPLY A BY 3 GIVING R ROUNDED.
    STOP RUN.

