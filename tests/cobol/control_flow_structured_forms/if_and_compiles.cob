*> vybe-test: cobol/control_flow_structured_forms/if_and_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF A = 1 AND B = 1 DISPLAY "Y" END-IF.
    STOP RUN.

