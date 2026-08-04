*> vybe-test: cobol/control_flow_structured_forms/alter_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    ALTER L1 TO PROCEED TO L2.
    GO TO L1.
L1. DISPLAY "A".
L2. DISPLAY "B".
    STOP RUN.

