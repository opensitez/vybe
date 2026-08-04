*> vybe-test: cobol/final_features/perform_test_after_runs_once
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 10.
PROCEDURE DIVISION.
    PERFORM WITH TEST AFTER UNTIL WS-I >= 5
        DISPLAY "Ran at least once"
        ADD 1 TO WS-I
    END-PERFORM.
    STOP RUN.

