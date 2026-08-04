*> vybe-test: cobol/new_features/is_numeric
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "12345".
PROCEDURE DIVISION.
    IF X IS NUMERIC
        DISPLAY "Number"
    END-IF.
    STOP RUN.

