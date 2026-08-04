*> vybe-test: cobol/category_table_handling/test_table_relative_indexing
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-REL-IDX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           SET IDX TO 2.
           MOVE 5 TO ELEM(IDX + 1).
           MOVE 4 TO ELEM(IDX - 1).
           DISPLAY ELEM(1) ELEM(3).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(1) DELIMITED SIZE ELEM(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "45"
        DISPLAY "FAIL: want [45] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

