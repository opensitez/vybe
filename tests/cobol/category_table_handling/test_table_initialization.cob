*> vybe-test: cobol/category_table_handling/test_table_initialization
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-INIT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR.
          05 ELEM OCCURS 3 TIMES PIC X VALUE "*".
       PROCEDURE DIVISION.
           DISPLAY ARR.
    MOVE SPACES TO WS-VYBE-L
    STRING ARR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "***"
        DISPLAY "FAIL: want [***] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

