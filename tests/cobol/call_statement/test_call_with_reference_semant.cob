*> vybe-test: cobol/call_statement/test_call_with_reference_semantics_runtime
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 1.
PROCEDURE DIVISION.

    CALL "SUBPROG" USING BY VALUE WS-A
        ON EXCEPTION
            DISPLAY WS-A
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
    STOP RUN.

