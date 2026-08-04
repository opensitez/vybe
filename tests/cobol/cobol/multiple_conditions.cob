*> vybe-test: cobol/cobol/multiple_conditions
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MULTICOND.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 5.
01 WS-B PIC 9(3) VALUE 10.
01 WS-C PIC 9(3) VALUE 15.
PROCEDURE DIVISION.
    IF WS-A < WS-B AND WS-B < WS-C
        DISPLAY "Ascending"
    END-IF.
    IF WS-A = 5 OR WS-B = 5
        DISPLAY "One is five"
    END-IF.
    IF NOT WS-A = 0
        DISPLAY "A is not zero"
    END-IF.
    STOP RUN.

