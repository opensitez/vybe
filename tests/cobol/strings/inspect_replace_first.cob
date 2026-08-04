*> vybe-test: cobol/strings/inspect_replace_first
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "aabaa".
PROCEDURE DIVISION.
    INSPECT TXT REPLACING FIRST "a" BY "X".
    STOP RUN.

