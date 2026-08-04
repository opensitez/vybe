*> vybe-test: cobol/arrays_tables_indexing/occurs_group_table_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TBL.
   05 ITM OCCURS 3 TIMES.
      10 V PIC X(3).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "AAA" TO V(1).
    MOVE "BBB" TO V(2).
    DISPLAY V(1).
    MOVE SPACES TO WS-VYBE-L
    STRING V(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAA"
        DISPLAY "FAIL: want [AAA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY V(2).
    MOVE SPACES TO WS-VYBE-L
    STRING V(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BBB"
        DISPLAY "FAIL: want [BBB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

