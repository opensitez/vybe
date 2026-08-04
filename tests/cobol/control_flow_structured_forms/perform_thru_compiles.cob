*> vybe-test: cobol/control_flow_structured_forms/perform_thru_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM P1 THRU P2.
    STOP RUN.
P1.
    DISPLAY "1".
P2.
    DISPLAY "2".

