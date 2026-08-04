*> vybe-test: cobol/call_statement/test_call_with_reference_only_on_exception
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(5) VALUE "HELLO".
PROCEDURE DIVISION.

    CALL "MISSING-PROG" USING BY REFERENCE WS-TEXT
        ON EXCEPTION
            DISPLAY "EX"
    END-CALL.
    STOP RUN.

