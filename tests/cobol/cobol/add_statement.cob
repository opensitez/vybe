*> vybe-test: cobol/cobol/add_statement
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. ARITH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 10.
01 WS-B PIC 9(5) VALUE 20.
01 WS-C PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    ADD WS-A TO WS-B.
    ADD WS-A WS-B GIVING WS-C.
    STOP RUN.

