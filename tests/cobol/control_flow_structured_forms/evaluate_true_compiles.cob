*> vybe-test: cobol/control_flow_structured_forms/evaluate_true_compiles
*> origin: languages/cobol/tests/cobol/test_control_flow_structured_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(2) VALUE 80.
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN N >= 90 DISPLAY "A"
        WHEN N >= 80 DISPLAY "B"
        WHEN OTHER DISPLAY "C"
    END-EVALUATE.
    STOP RUN.

