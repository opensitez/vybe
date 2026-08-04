*> vybe-test: cobol/control_flow/eval_true
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 85.
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN X >= 90
            DISPLAY "A"
        WHEN X >= 80
            DISPLAY "B"
        WHEN X >= 70
            DISPLAY "C"
        WHEN OTHER
            DISPLAY "F"
    END-EVALUATE.
    STOP RUN.

