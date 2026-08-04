*> vybe-test: cobol/exception_handling/raise_exception_followed_by_branch_compile
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ERR PIC X(4) VALUE "FAIL".
PROCEDURE DIVISION.
    IF WS-ERR = "FAIL"
        RAISE EXCEPTION "boom"
    END-IF.
    STOP RUN.

