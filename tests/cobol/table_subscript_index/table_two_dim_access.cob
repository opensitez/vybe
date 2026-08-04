*> vybe-test: cobol/table_subscript_index/table_two_dim_access
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 M.
   05 ROW OCCURS 3 TIMES.
      10 COL PIC 9 OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 5 TO COL(2 2).
    DISPLAY COL(2 2).
    MOVE SPACES TO WS-VYBE-L
    STRING COL(2 DELIMITED SIZE 2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

