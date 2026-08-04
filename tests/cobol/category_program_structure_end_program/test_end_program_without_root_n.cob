*> vybe-test: cobol/category_program_structure_end_program/test_end_program_without_root_name_token
*> origin: languages/cobol/tests/cobol/test_category_program_structure_end_program.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'ROOT'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'ROOT' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ROOT"
        DISPLAY "FAIL: want [ROOT] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN. END PROGRAM

