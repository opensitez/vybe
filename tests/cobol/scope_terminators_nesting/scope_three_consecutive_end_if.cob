*> vybe-test: cobol/scope_terminators_nesting/scope_three_consecutive_end_if
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
01 C PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A = 1
        DISPLAY "A"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    IF B = 2
        DISPLAY "B"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    IF C = 3
        DISPLAY "C"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "C" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "C"
        DISPLAY "FAIL: want [C] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

