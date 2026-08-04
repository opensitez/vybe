*> vybe-test: cobol/enterprise/func_chain
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "  Hello  ".
01 WS-OUT PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(FUNCTION TRIM(WS-TEXT)) TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.

