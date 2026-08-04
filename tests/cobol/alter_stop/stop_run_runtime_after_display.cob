*> vybe-test: cobol/alter_stop/stop_run_runtime_after_display
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           DISPLAY "start"
           STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING "start" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "start"
        DISPLAY "FAIL: want [start] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           DISPLAY "after"

