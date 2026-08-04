*> vybe-test: cobol/exception_handling/call_end_call_exception_branches_compile
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "WORK"
        ON EXCEPTION DISPLAY "ERR"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

