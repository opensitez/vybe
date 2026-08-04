*> vybe-test: cobol/cobol/compute_statement
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. COMP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 10.
01 WS-B PIC 9(5) VALUE 3.
01 WS-RESULT PIC 9(10) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-RESULT = WS-A + WS-B * 2.
    COMPUTE WS-RESULT = (WS-A + WS-B) * 2.
    COMPUTE WS-RESULT = WS-A ** 2.
    STOP RUN.

