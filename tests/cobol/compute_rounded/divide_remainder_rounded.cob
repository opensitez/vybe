*> vybe-test: cobol/compute_rounded/divide_remainder_rounded
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 Q PIC 9(4) VALUE 0.
01 REM PIC 9(4)V9 VALUE 0.
PROCEDURE DIVISION.
    DIVIDE 7 INTO 22 GIVING Q REMAINDER REM.
    STOP RUN.

