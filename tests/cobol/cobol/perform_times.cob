*> vybe-test: cobol/cobol/perform_times
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PTIMES.
PROCEDURE DIVISION.
    PERFORM 5 TIMES
        DISPLAY "Hello"
    END-PERFORM.
    STOP RUN.

