*> vybe-test: cobol/compute_rounded/subtract_giving_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(4)V9 VALUE 100.5.
01 B PIC 9(4)V9 VALUE 33.3.
01 R PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    SUBTRACT B FROM A GIVING R ROUNDED.
    STOP RUN.

