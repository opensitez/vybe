*> vybe-test: cobol/multiply_advanced/test_multiply_negative
*> origin: languages/cobol/tests/cobol/test_multiply_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

01 WS-A PIC S9(3) VALUE 5.
01 WS-B PIC S9(3) VALUE -2.
01 WS-C PIC S9(3) VALUE 0.
    STOP RUN.

