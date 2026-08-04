*> vybe-test: cobol/exception_handling/raise_named_exception_with_id_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10) VALUE "EX-1".
PROCEDURE DIVISION.
    IF WS-NAME = "EX-1"
        RAISE EXCEPTION EC-IMPLICIT-EXCEPTION
    END-IF.
    STOP RUN.

