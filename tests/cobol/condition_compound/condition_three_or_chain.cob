*> vybe-test: cobol/condition_compound/condition_three_or_chain
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC X VALUE "C".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF C = "A" OR C = "B" OR C = "C"
        DISPLAY "ABC"
    ELSE
        DISPLAY "OTHER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ABC" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC"
        DISPLAY "FAIL: want [ABC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

