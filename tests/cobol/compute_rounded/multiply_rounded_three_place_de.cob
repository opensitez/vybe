*> vybe-test: cobol/compute_rounded/multiply_rounded_three_place_decimal
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V999 VALUE 1.333.
01 R PIC 9(4)V99 VALUE 0.
PROCEDURE DIVISION.
    MULTIPLY 3 BY A GIVING R ROUNDED.
    STOP RUN.

