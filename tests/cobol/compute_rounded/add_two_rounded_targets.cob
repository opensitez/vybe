*> vybe-test: cobol/compute_rounded/add_two_rounded_targets
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V9 VALUE 1.5.
01 B PIC 9(3)V9 VALUE 2.5.
01 R1 PIC 9(3) VALUE 0.
01 R2 PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    ADD A B GIVING R1 ROUNDED R2 ROUNDED.
    STOP RUN.

