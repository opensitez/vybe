*> vybe-test: cobol/structured_control_flow_expanded/perform_thru_with_exit_paragraph_compiles
*> origin: languages/cobol/tests/cobol/test_structured_control_flow_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FLOW-A.
PROCEDURE DIVISION.
    PERFORM STEP-A THRU STEP-C.
    STOP RUN.
STEP-A.
    DISPLAY "A".
STEP-B.
    DISPLAY "B".
STEP-C.
    DISPLAY "C".

