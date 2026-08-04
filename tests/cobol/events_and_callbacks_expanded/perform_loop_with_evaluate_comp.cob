*> vybe-test: cobol/events_and_callbacks_expanded/perform_loop_with_evaluate_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL N >= 2
        ADD 1 TO N
        EVALUATE N
            WHEN 1 DISPLAY "A"
            WHEN 2 DISPLAY "B"
        END-EVALUATE
    END-PERFORM.
    STOP RUN.

