*> vybe-test: cobol/new_features/is_alphabetic_lower
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "hello".
PROCEDURE DIVISION.
    IF X IS ALPHABETIC-LOWER
        DISPLAY "Lower"
    END-IF.
    STOP RUN.

