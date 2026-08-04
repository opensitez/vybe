*> vybe-test: cobol/condition_compound/condition_not_alphabetic
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "12345".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT S IS ALPHABETIC
        DISPLAY "NOT ALPHA"
    ELSE
        DISPLAY "ALPHA"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "NOT ALPHA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NOT ALPHA"
        DISPLAY "FAIL: want [NOT ALPHA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

