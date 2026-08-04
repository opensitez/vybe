*> vybe-test: cobol/control_flow_structured_forms/goto_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    GO TO L1.
L1.
    DISPLAY "OK".
    STOP RUN.

