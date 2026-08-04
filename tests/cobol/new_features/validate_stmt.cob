*> vybe-test: cobol/new_features/validate_stmt
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(5) VALUE 123.
PROCEDURE DIVISION.
    VALIDATE X.
    STOP RUN.

