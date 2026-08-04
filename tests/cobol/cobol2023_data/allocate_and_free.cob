*> vybe-test: cobol/cobol2023_data/allocate_and_free
*> origin: languages/cobol/tests/cobol/test_cobol2023_data.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR USAGE POINTER.
PROCEDURE DIVISION.
    ALLOCATE WS-PTR.
    FREE WS-PTR.
    STOP RUN.

