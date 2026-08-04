*> vybe-test: cobol/control_flow_structured_forms/perform_times_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 3 TIMES DISPLAY "L" END-PERFORM.
    STOP RUN.

