*> vybe-test: cobol/new_features/inspect_converting_spaces
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "a b c".
PROCEDURE DIVISION.
    INSPECT TXT CONVERTING " " TO "-".
    STOP RUN.

