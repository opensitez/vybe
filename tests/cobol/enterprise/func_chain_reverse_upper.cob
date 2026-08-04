*> vybe-test: cobol/enterprise/func_chain_reverse_upper
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(10) VALUE "hello".
01 WS-OUT PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION REVERSE(FUNCTION UPPER-CASE(WS-TEXT)) TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.

