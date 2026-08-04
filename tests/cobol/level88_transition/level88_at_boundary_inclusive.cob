*> vybe-test: cobol/level88_transition/level88_at_boundary_inclusive
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TEMP PIC S9(3) VALUE -5.
    88 FREEZING VALUE -50 THRU 0.
PROCEDURE DIVISION.
    IF FREEZING
        DISPLAY "COLD"
    END-IF.
    STOP RUN.

