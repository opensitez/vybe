*> vybe-test: cobol/cobol/display_multiple_runtime_single_line
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. DISPMULRUN.
PROCEDURE DIVISION.
    DISPLAY "Name: " "Alice" " Age: " 30.
    MOVE SPACES TO WS-VYBE-L
    STRING "Name: " DELIMITED SIZE "Alice" DELIMITED SIZE " Age: " DELIMITED SIZE 30 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Name: Alice Age: 30"
        DISPLAY "FAIL: want [Name: Alice Age: 30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

