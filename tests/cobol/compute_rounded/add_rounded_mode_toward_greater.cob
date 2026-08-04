*> vybe-test: cobol/compute_rounded/add_rounded_mode_toward_greater
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V9 VALUE 5.3.
01 R PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    ADD A TO R ROUNDED MODE TOWARD-GREATER.
    STOP RUN.

