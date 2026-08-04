*> vybe-test: cobol/new_features/is_not_numeric
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "Hello".
PROCEDURE DIVISION.
    IF X IS NOT NUMERIC
        DISPLAY "Not a number"
    END-IF.
    STOP RUN.

