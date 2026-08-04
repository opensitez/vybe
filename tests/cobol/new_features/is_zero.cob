*> vybe-test: cobol/new_features/is_zero
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    IF X IS ZERO
        DISPLAY "Zero"
    END-IF.
    STOP RUN.

