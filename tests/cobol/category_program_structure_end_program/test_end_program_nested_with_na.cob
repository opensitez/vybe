*> vybe-test: cobol/category_program_structure_end_program/test_end_program_nested_with_named_terminator
*> origin: languages/cobol/tests/cobol/test_category_program_structure_end_program.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. CALL 'S1'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. S1. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. EXIT PROGRAM. END PROGRAM S1. END PROGRAM T.

