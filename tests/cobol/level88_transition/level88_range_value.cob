*> vybe-test: cobol/level88_transition/level88_range_value
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SCORE PIC 9(3) VALUE 75.
    88 PASSING-SCORE VALUE 60 THRU 100.
PROCEDURE DIVISION.
    IF PASSING-SCORE
        DISPLAY "PASS"
    END-IF.
    STOP RUN.

