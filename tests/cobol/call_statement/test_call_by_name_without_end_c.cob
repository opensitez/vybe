*> vybe-test: cobol/call_statement/test_call_by_name_without_end_call_compiles
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUBPROG".
    STOP RUN.

