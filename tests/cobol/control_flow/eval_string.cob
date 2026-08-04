*> vybe-test: cobol/control_flow/eval_string
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(5) VALUE "B".
PROCEDURE DIVISION.
    EVALUATE X
        WHEN "A"
            DISPLAY "Alpha"
        WHEN "B"
            DISPLAY "Beta"
    END-EVALUATE.
    STOP RUN.

