*> vybe-test: cobol/strings/inspect_tally_all
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "Hello World".
01 CNT PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT TXT TALLYING CNT FOR ALL "l".
    STOP RUN.

