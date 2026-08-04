*> vybe-test: cobol/compute_rounded/subtract_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V9 VALUE 10.5.
01 R PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    SUBTRACT 0.5 FROM A ROUNDED.
    STOP RUN.

