*> vybe-test: cobol/control_flow/perform_para
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM MY-PARA.
    STOP RUN.
MY-PARA.
    DISPLAY "Hello".

