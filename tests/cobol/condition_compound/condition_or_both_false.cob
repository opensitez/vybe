*> vybe-test: cobol/condition_compound/condition_or_both_false
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
01 B PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A > 0 OR B > 0
        DISPLAY "EITHER"
    ELSE
        DISPLAY "NEITHER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "EITHER" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "EITHER"
        DISPLAY "FAIL: want [EITHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

