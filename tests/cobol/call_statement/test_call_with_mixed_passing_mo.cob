*> vybe-test: cobol/call_statement/test_call_with_mixed_passing_modes
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC 9(3) VALUE 10.
01 WS-TEXT PIC X(4) VALUE "ABCD".
PROCEDURE DIVISION.

    CALL "SUBPROG"
        USING BY VALUE WS-ARG
        BY REFERENCE WS-TEXT
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
    STOP RUN.

