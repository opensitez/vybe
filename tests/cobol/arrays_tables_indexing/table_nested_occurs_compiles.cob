*> vybe-test: cobol/arrays_tables_indexing/table_nested_occurs_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 O1 OCCURS 2 TIMES.
      10 O2 OCCURS 2 TIMES.
         15 V PIC X(2).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "AA" TO V(1,1).
    MOVE "BB" TO V(2,2).
    DISPLAY V(1,1).
    MOVE SPACES TO WS-VYBE-L
    STRING V(1,1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AA"
        DISPLAY "FAIL: want [AA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY V(2,2).
    MOVE SPACES TO WS-VYBE-L
    STRING V(2,2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BB"
        DISPLAY "FAIL: want [BB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

