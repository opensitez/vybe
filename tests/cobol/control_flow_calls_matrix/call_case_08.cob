*> vybe-test: cobol/control_flow_calls_matrix/call_case_08
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUB8" ON EXCEPTION DISPLAY "E" NOT ON EXCEPTION DISPLAY "O" END-CALL.
    STOP RUN.

