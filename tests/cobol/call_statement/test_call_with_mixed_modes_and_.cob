*> vybe-test: cobol/call_statement/test_call_with_mixed_modes_and_on_exception
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC X(5) VALUE "HELLO".
01 WS-NUM PIC 9(3) VALUE 123.
01 WS-PROG PIC X(20) VALUE "SUBPROG".
PROCEDURE DIVISION.

    CALL WS-PROG
        USING BY REFERENCE WS-ARG
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
        END-CALL.
    STOP RUN.

