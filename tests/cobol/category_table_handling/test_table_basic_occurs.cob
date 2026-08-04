*> vybe-test: cobol/category_table_handling/test_table_basic_occurs
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-OCCURS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR.
          05 ELEM OCCURS 5 TIMES PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 7 TO ELEM(3).
           DISPLAY ELEM(3).
    MOVE SPACES TO WS-VYBE-L
    STRING ELEM(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7"
        DISPLAY "FAIL: want [7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

