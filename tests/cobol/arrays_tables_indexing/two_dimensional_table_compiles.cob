*> vybe-test: cobol/arrays_tables_indexing/two_dimensional_table_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 R OCCURS 2 TIMES.
      10 C PIC 9 OCCURS 2 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 5 TO C(1,1).
    MOVE 6 TO C(1,2).
    MOVE 7 TO C(2,1).
    MOVE 8 TO C(2,2).
    DISPLAY C(2,2).
    MOVE SPACES TO WS-VYBE-L
    STRING C(2,2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "8"
        DISPLAY "FAIL: want [8] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY C(1,2).
    MOVE SPACES TO WS-VYBE-L
    STRING C(1,2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "6"
        DISPLAY "FAIL: want [6] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

