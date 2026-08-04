*> vybe-test: cobol/control_flow_structured_forms/continue_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF X = 1 CONTINUE END-IF.
    STOP RUN.

