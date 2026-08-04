*> vybe-test: cobol/new_features/inspect_converting
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "Hello World".
PROCEDURE DIVISION.
    INSPECT TXT CONVERTING "abcdefghij" TO "ABCDEFGHIJ".
    STOP RUN.

