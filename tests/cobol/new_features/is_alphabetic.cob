*> vybe-test: cobol/new_features/is_alphabetic
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "Hello".
PROCEDURE DIVISION.
    IF X IS ALPHABETIC
        DISPLAY "Alpha"
    END-IF.
    STOP RUN.

