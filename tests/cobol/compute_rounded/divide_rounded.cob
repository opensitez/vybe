*> vybe-test: cobol/compute_rounded/divide_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 10.
01 R PIC 9(3)V9 VALUE 0.
PROCEDURE DIVISION.
    DIVIDE 3 INTO A GIVING R ROUNDED.
    STOP RUN.

