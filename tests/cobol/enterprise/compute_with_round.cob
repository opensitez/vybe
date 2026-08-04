*> vybe-test: cobol/enterprise/compute_with_round
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-R PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-R = 10 / 3.
    STOP RUN.

