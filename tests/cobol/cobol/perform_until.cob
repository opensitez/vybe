*> vybe-test: cobol/cobol/perform_until
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PUNTIL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    PERFORM UNTIL WS-I > 10
        DISPLAY WS-I
        ADD 1 TO WS-I
    END-PERFORM.
    STOP RUN.

