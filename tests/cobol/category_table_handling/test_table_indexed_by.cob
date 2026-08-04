*> vybe-test: cobol/category_table_handling/test_table_indexed_by
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-INDEX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC X(2).
       PROCEDURE DIVISION.
           SET IDX TO 2.
           MOVE "AB" TO ELEM(IDX).
           DISPLAY ELEM(IDX).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(IDX) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AB"
        DISPLAY "FAIL: want [AB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

