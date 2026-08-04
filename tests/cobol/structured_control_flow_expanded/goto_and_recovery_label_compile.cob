*> vybe-test: cobol/structured_control_flow_expanded/goto_and_recovery_label_compiles
*> origin: languages/cobol/tests/cobol/test_structured_control_flow_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FLOW-B.
PROCEDURE DIVISION.
    GO TO RECOVER-LABEL.
MAIN-LABEL.
    DISPLAY "MAIN".
    STOP RUN.
RECOVER-LABEL.
    DISPLAY "RECOVER".
    GO TO MAIN-LABEL.

