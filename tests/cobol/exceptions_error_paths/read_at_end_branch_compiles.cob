*> vybe-test: cobol/exceptions_error_paths/read_at_end_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X(80).
PROCEDURE DIVISION.
    READ WS-FILE
        AT END DISPLAY "EOF"
    END-READ.
    STOP RUN.

