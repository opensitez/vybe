*> vybe-test: cobol/enterprise/usage_comp3
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(9) COMP-3.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

