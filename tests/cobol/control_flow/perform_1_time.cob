*> vybe-test: cobol/control_flow/perform_1_time
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 1 TIMES
        DISPLAY "Once"
    END-PERFORM.
    STOP RUN.

