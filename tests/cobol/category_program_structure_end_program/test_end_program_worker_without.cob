*> vybe-test: cobol/category_program_structure_end_program/test_end_program_worker_without_local_end_name
*> origin: languages/cobol/tests/cobol/test_category_program_structure_end_program.rs
IDENTIFICATION DIVISION. PROGRAM-ID. ROOT. PROCEDURE DIVISION. CALL 'WORKER'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. WORKER. PROCEDURE DIVISION. DISPLAY 'WORKER'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'WORKER' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "WORKER"
        DISPLAY "FAIL: want [WORKER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END PROGRAM.

