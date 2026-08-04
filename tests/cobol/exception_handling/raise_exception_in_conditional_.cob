*> vybe-test: cobol/exception_handling/raise_exception_in_conditional_flow_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF FLAG = 1
        RAISE EXCEPTION "E1"
    ELSE
        DISPLAY "NOERR"
    END-IF.
    STOP RUN.

