*> vybe-test: cobol/new_features/is_alphabetic_upper
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "HELLO".
PROCEDURE DIVISION.
    IF X IS ALPHABETIC-UPPER
        DISPLAY "Upper"
    END-IF.
    STOP RUN.

