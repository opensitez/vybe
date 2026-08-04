*> vybe-test: cobol/new_features/rewrite_basic
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC PIC X(80).
PROCEDURE DIVISION.
    REWRITE REC.
    STOP RUN.

