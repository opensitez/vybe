*> vybe-test: cobol/call_statement/test_call_using_parameters
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 100.
01 WS-B PIC X(5) VALUE "HELLO".
PROCEDURE DIVISION.

    CALL "SUBPROG" USING BY REFERENCE WS-A
                         BY CONTENT WS-B.
    CALL "SUBPROG" USING BY VALUE WS-A.
    STOP RUN.

