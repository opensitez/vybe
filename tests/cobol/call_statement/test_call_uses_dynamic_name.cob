*> vybe-test: cobol/call_statement/test_call_uses_dynamic_name
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PROG PIC X(20) VALUE "SUBPROG".
PROCEDURE DIVISION.

    CALL WS-PROG
        ON EXCEPTION DISPLAY "MISS"
        NOT ON EXCEPTION DISPLAY "HIT"
    END-CALL.
    STOP RUN.

