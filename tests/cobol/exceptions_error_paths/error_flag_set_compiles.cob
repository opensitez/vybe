*> vybe-test: cobol/exceptions_error_paths/error_flag_set_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ERR-FLAG PIC 9 VALUE 0.
PROCEDURE DIVISION.
    CALL "MAYBE"
        ON EXCEPTION MOVE 1 TO ERR-FLAG
    END-CALL.
    STOP RUN.

