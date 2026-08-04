*> vybe-test: cobol/condition_compound/condition_and_one_false
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A > 0 AND B > 0
        DISPLAY "BOTH"
    ELSE
        DISPLAY "NOT BOTH"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BOTH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NOT BOTH"
        DISPLAY "FAIL: want [NOT BOTH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

