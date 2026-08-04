*> vybe-test: cobol/category_table_handling/test_table_occurs_depending_on
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-ODO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ARR-LEN PIC 9 VALUE 2.
       01 ARR.
          05 ELEM OCCURS 1 TO 5 TIMES DEPENDING ON ARR-LEN PIC X.
       PROCEDURE DIVISION.
           MOVE "A" TO ELEM(1).
           MOVE "B" TO ELEM(2).
           DISPLAY ARR.
    MOVE SPACES TO WS-VYBE-L
    STRING ARR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AB"
        DISPLAY "FAIL: want [AB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

