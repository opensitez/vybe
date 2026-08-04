*> vybe-test: cobol/cobol/multiply_statement
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MUL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 5.
01 WS-B PIC 9(5) VALUE 3.
01 WS-C PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    MULTIPLY WS-A BY WS-B.
    MULTIPLY WS-A BY WS-B GIVING WS-C.
    STOP RUN.

