*> vybe-test: cobol/exception_handling/call_with_not_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "MAY-FAIL"
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
        END-CALL.
    STOP RUN.

