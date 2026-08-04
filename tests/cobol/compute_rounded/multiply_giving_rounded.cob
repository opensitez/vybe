*> vybe-test: cobol/compute_rounded/multiply_giving_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V99 VALUE 1.25.
01 B PIC 9(2) VALUE 3.
01 R PIC 9(3)V9 VALUE 0.
PROCEDURE DIVISION.
    MULTIPLY A BY B GIVING R ROUNDED.
    STOP RUN.

