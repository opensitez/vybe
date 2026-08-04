*> vybe-test: cobol/enterprise/function_chaining_program
*> origin: languages/cobol/tests/cobol/test_enterprise.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. FCHAIN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT  PIC X(50) VALUE "  hello, world!  ".
01 WS-OUTPUT PIC X(50).
01 WS-LEN    PIC 9(5).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(FUNCTION TRIM(WS-INPUT))
         TO WS-OUTPUT.
    MOVE FUNCTION LENGTH(FUNCTION TRIM(WS-INPUT))
         TO WS-LEN.
    DISPLAY "Result: " WS-OUTPUT.
    DISPLAY "Length: " WS-LEN.
    STOP RUN.

