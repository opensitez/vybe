*> vybe-test: cobol/control_flow_structured_forms/perform_paragraph_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM P1.
    STOP RUN.
P1.
    DISPLAY "P".

