*> vybe-test: cobol/category_program_structure_end_program/test_end_program_nested_without_main_terminator
*> origin: languages/cobol/tests/cobol/test_category_program_structure_end_program.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. CALL 'S1'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. S1. PROCEDURE DIVISION. DISPLAY 'SUB'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'SUB' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SUB"
        DISPLAY "FAIL: want [SUB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. EXIT PROGRAM. END PROGRAM.

