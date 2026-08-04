*> vybe-test: cobol/enterprise/usage_binary
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(9) USAGE BINARY.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

