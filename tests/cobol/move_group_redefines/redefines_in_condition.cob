*> vybe-test: cobol/move_group_redefines/redefines_in_condition
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 STATUS-BYTE PIC X VALUE "Y".
01 STATUS-NUM REDEFINES STATUS-BYTE PIC 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF STATUS-BYTE = "Y"
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

