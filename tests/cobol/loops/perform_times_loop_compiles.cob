*> vybe-test: cobol/loops/perform_times_loop_compiles
*> origin: languages/cobol/tests/cobol/test_loops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 3 TIMES
        DISPLAY "loop"
    END-PERFORM.
    STOP RUN.

