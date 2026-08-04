*> vybe-test: cobol/cobol/continue_stmt
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CONT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF WS-X > 10
        DISPLAY "Big"
    ELSE
        CONTINUE
    END-IF.
    STOP RUN.

