*> vybe-test: cobol/control_flow/perform_thru
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM INIT-PARA THRU CLEANUP-PARA.
    STOP RUN.
INIT-PARA.
    DISPLAY "Init".
CLEANUP-PARA.
    DISPLAY "Cleanup".

