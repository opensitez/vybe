*> vybe-test: cobol/category_program_structure/test_prog_nested_common
*> origin: languages/cobol/tests/cobol/test_category_program_structure.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. CALL 'S1'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. S1 IS COMMON PROGRAM. PROCEDURE DIVISION. DISPLAY '1'.
    MOVE SPACES TO WS-VYBE-L
    STRING '1' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. EXIT PROGRAM. END PROGRAM S1.

