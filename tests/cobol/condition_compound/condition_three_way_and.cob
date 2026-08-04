*> vybe-test: cobol/condition_compound/condition_three_way_and
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
01 C PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A < B AND B < C AND A < C
        DISPLAY "ORDERED"
    ELSE
        DISPLAY "NOT ORDERED"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ORDERED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ORDERED"
        DISPLAY "FAIL: want [ORDERED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

