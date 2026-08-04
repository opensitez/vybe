*> vybe-test: cobol/nested_if_else/if_boundary_exclusive_gt
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(2) VALUE 10.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 10
        DISPLAY "GT"
    ELSE
        DISPLAY "LE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "GT" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LE"
        DISPLAY "FAIL: want [LE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

