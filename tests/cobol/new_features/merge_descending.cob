*> vybe-test: cobol/new_features/merge_descending
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    MERGE WS-FILE ON DESCENDING KEY WS-KEY.
    STOP RUN.

