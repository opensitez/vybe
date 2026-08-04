*> vybe-test: cobol/new_features/rewrite_from
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC PIC X(80).
01 NEW-REC PIC X(80) VALUE "Updated".
PROCEDURE DIVISION.
    REWRITE REC FROM NEW-REC.
    STOP RUN.

