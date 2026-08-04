*> vybe-test: cobol/new_features/is_negative
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC S9(5) VALUE -5.
PROCEDURE DIVISION.
    IF X IS NEGATIVE
        DISPLAY "Negative"
    END-IF.
    STOP RUN.

