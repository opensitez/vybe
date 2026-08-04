*> vybe-test: cobol/exceptions_error_paths/write_invalid_key_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X(80).
PROCEDURE DIVISION.
    WRITE F
        INVALID KEY DISPLAY "BAD"
    END-WRITE.
    STOP RUN.

