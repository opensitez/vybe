*> vybe-test: cobol/control_flow/perform_nested_times
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 3 TIMES
        PERFORM 2 TIMES
            DISPLAY "Inner"
        END-PERFORM
    END-PERFORM.
    STOP RUN.

