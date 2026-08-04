*> vybe-test: cobol/arithmetic_control_flow/call_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUB-X"
        ON EXCEPTION DISPLAY "ERR"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

