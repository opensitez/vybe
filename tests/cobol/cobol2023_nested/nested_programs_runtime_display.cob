*> vybe-test: cobol/cobol2023_nested/nested_programs_runtime_display
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG.
PROCEDURE DIVISION.
    DISPLAY "Main program".
    MOVE SPACES TO WS-VYBE-L
    STRING "Main program" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Main program"
        DISPLAY "FAIL: want [Main program] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    END PROGRAM MAIN-PROG.

