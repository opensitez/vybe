*> vybe-test: cobol/exception_handling/call_with_exception_path_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "SUB".
    DISPLAY "done".
    STOP RUN.

