*> vybe-test: cobol/new_features/start_file
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    START WS-FILE KEY = WS-KEY.
    STOP RUN.

