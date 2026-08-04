*> vybe-test: cobol/control_flow/eval_no_other
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    EVALUATE X
        WHEN 1
            DISPLAY "One"
        WHEN 2
            DISPLAY "Two"
    END-EVALUATE.
    STOP RUN.

