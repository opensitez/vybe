*> vybe-test: cobol/arrays_tables_indexing/table_group_move_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G1.
   05 A PIC X(2) OCCURS 2 TIMES.
01 G2.
   05 B PIC X(2) OCCURS 2 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "AA" TO A(1).
    MOVE "BB" TO A(2).
    MOVE G1 TO G2.
    DISPLAY B(1).
    MOVE SPACES TO WS-VYBE-L
    STRING B(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AA"
        DISPLAY "FAIL: want [AA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY B(2).
    MOVE SPACES TO WS-VYBE-L
    STRING B(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BB"
        DISPLAY "FAIL: want [BB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

