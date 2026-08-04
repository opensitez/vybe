*> vybe-test: cobol/exception_handling/call_with_exception_and_recovery_path_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(2) VALUE "OK".
PROCEDURE DIVISION.
    CALL "MAY-FAIL"
        ON EXCEPTION MOVE "ER" TO WS-STATUS
    END-CALL.
    IF WS-STATUS = "ER" DISPLAY "RECOVER" END-IF.
    STOP RUN.

