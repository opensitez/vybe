*> vybe-test: cobol/category_table_handling/test_table_indexed_set_up_down_paths
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SETPATH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           MOVE 1 TO ELEM(IDX).
           SET IDX UP BY 2.
           MOVE 3 TO ELEM(IDX).
           SET IDX DOWN BY 1.
           DISPLAY ELEM(1).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           DISPLAY ELEM(2).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           DISPLAY ELEM(3).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

