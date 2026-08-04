*> vybe-test: cobol/control_flow_structured_forms/if_not_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
PROCEDURE DIVISION.
    IF NOT A = 1 DISPLAY "Y" END-IF.
    STOP RUN.

