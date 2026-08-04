*> vybe-test: cobol/final_features/copy_replacing
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING OLD-NAME BY NEW-NAME.
    STOP RUN.

