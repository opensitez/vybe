*> vybe-test: cobol/scope_terminators/test_terminator_period_closing
*> origin: languages/cobol/tests/cobol/test_scope_terminators.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 5.
PROCEDURE DIVISION.

    IF WS-A > 0
        DISPLAY "POS"
        IF WS-A = 5
            DISPLAY "FIVE".
    STOP RUN.

