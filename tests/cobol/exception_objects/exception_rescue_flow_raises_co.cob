*> vybe-test: cobol/exception_objects/exception_rescue_flow_raises_compiles
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. test.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ERR PIC X VALUE "N".
PROCEDURE DIVISION.
    CALL "NO-SUCH-ROUTINE"
        ON EXCEPTION
            MOVE "E" TO WS-ERR
    END-CALL
    DISPLAY WS-ERR
    STOP RUN.

