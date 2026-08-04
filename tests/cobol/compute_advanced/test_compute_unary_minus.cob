*> vybe-test: cobol/compute_advanced/test_compute_unary_minus
*> origin: languages/cobol/tests/cobol/test_compute_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

01 WS-A PIC S9(3) VALUE 42.
01 WS-B PIC S9(3) VALUE 0.
    STOP RUN.

