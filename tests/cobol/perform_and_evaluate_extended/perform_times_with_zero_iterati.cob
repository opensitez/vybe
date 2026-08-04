*> vybe-test: cobol/perform_and_evaluate_extended/perform_times_with_zero_iterations_is_accepted
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9 VALUE 0.
PROCEDURE DIVISION.

    PERFORM 0 TIMES
        ADD 1 TO WS-I
    END-PERFORM.
    STOP RUN.

