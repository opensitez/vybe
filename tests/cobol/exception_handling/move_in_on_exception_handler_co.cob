*> vybe-test: cobol/exception_handling/move_in_on_exception_handler_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-OUT PIC X(6) VALUE SPACES.
PROCEDURE DIVISION.
    CALL "MAY-FAIL"
        ON EXCEPTION
            MOVE "FAILED" TO WS-OUT
        END-CALL
    DISPLAY WS-OUT.
    STOP RUN.

