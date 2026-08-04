*> vybe-test: cobol/control_flow_structured_forms/if_else_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 2.
PROCEDURE DIVISION.
    IF X = 1 DISPLAY "A" ELSE DISPLAY "B" END-IF.
    STOP RUN.

