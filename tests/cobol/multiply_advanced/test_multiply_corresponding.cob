*> vybe-test: cobol/multiply_advanced/test_multiply_corresponding
*> origin: languages/cobol/tests/cobol/test_multiply_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-G1.
   05 WS-X PIC 9(3) VALUE 2.
   05 WS-Y PIC 9(3) VALUE 3.
01 WS-G2.
   05 WS-X PIC 9(3) VALUE 10.
   05 WS-Y PIC 9(3) VALUE 20.
PROCEDURE DIVISION.

    MULTIPLY CORRESPONDING WS-G1 BY WS-G2.
    STOP RUN.

