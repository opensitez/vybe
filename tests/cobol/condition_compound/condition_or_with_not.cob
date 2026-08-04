*> vybe-test: cobol/condition_compound/condition_or_with_not
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
01 B PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT A > 0 OR NOT B > 0
        DISPLAY "YES"
    ELSE
        DISPLAY "NO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "YES"
        DISPLAY "FAIL: want [YES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

