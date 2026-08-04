*> vybe-test: cobol/category_program_structure_end_program/test_end_program_main_omitted_name
*> origin: languages/cobol/tests/cobol/test_category_program_structure_end_program.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'MAIN'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'MAIN' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MAIN"
        DISPLAY "FAIL: want [MAIN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN. END PROGRAM.

