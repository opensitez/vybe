*> vybe-test: cobol/control_flow/perform_multiple_paras
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM STEP1-PARA.
    PERFORM STEP2-PARA.
    PERFORM STEP3-PARA.
    STOP RUN.
STEP1-PARA.
    DISPLAY "Step 1".
STEP2-PARA.
    DISPLAY "Step 2".
STEP3-PARA.
    DISPLAY "Step 3".

