*> vybe-test: cobol/category_table_handling/test_table_multidimensional
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 MATRIX.
          05 ROW OCCURS 3 TIMES.
             10 COL OCCURS 3 TIMES PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 9 TO COL(2, 3).
           DISPLAY COL(2, 3).
    MOVE SPACES TO WS-VYBE-L
    STRING COL(2, DELIMITED SIZE 3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9"
        DISPLAY "FAIL: want [9] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

