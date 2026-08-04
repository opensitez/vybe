*> vybe-test: cobol/cobol2023_data/validate_statement
*> origin: languages/cobol/tests/cobol/test_cobol2023_data.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT PIC X(20).
PROCEDURE DIVISION.
    VALIDATE WS-INPUT.
    STOP RUN.

