*> vybe-test: cobol/cobol/raise_exception
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. EXCEPT.
PROCEDURE DIVISION.
    RAISE EXCEPTION "Something went wrong".
    STOP RUN.

