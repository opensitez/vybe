*> vybe-test: cobol/control_flow_structured_forms/evaluate_basic_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE X
        WHEN 1 DISPLAY "A"
        WHEN 2 DISPLAY "B"
        WHEN OTHER DISPLAY "C"
    END-EVALUATE.
    STOP RUN.

