*> vybe-test: cobol/strings/inspect_replace_all
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "Hello World".
PROCEDURE DIVISION.
    INSPECT TXT REPLACING ALL "l" BY "r".
    STOP RUN.

