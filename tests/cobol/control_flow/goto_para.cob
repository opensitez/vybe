*> vybe-test: cobol/control_flow/goto_para
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    DISPLAY "Start".
    STOP RUN.
ERROR-PARA.
    DISPLAY "Error".

