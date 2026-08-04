*> vybe-test: cobol/control_flow/perform_10_times
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 10 TIMES
        DISPLAY "Loop"
    END-PERFORM.
    STOP RUN.

