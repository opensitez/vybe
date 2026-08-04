*> vybe-test: cobol/new_features/is_positive
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC S9(5) VALUE 10.
PROCEDURE DIVISION.
    IF X IS POSITIVE
        DISPLAY "Positive"
    END-IF.
    STOP RUN.

