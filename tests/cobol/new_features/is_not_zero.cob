*> vybe-test: cobol/new_features/is_not_zero
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(5) VALUE 5.
PROCEDURE DIVISION.
    IF X IS NOT ZERO
        DISPLAY "Non-zero"
    END-IF.
    STOP RUN.

