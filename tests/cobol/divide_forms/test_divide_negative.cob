*> vybe-test: cobol/divide_forms/test_divide_negative
*> origin: languages/cobol/tests/cobol/test_divide_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

01 WS-A PIC S9(3) VALUE -10.
01 WS-B PIC S9(3) VALUE 2.
01 WS-C PIC S9(3) VALUE 0.
    STOP RUN.

