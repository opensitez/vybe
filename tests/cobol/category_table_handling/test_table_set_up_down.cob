*> vybe-test: cobol/category_table_handling/test_table_set_up_down
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SET-MATH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC 9.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SET IDX UP BY 2.
           SET IDX DOWN BY 1.
           MOVE 8 TO ELEM(IDX).
           DISPLAY ELEM(2).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "8"
        DISPLAY "FAIL: want [8] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

