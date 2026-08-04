*> vybe-test: cobol/category_table_handling/test_table_copy_between_tables
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-COPY.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 SRC.
          05 SRC-ELEM OCCURS 3 TIMES PIC X(2).
       01 DST.
          05 DST-ELEM OCCURS 3 TIMES PIC X(2).
       PROCEDURE DIVISION.
           MOVE "AA" TO SRC-ELEM(1).
           MOVE "BB" TO SRC-ELEM(2).
           MOVE "CC" TO SRC-ELEM(3).
           MOVE SRC TO DST.
           DISPLAY DST-ELEM(1).
    MOVE SPACES TO WS-VYBE-L
    STRING DST-ELEM(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AA"
        DISPLAY "FAIL: want [AA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           DISPLAY DST-ELEM(2).
    MOVE SPACES TO WS-VYBE-L
    STRING DST-ELEM(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BB"
        DISPLAY "FAIL: want [BB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           DISPLAY DST-ELEM(3).
    MOVE SPACES TO WS-VYBE-L
    STRING DST-ELEM(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CC"
        DISPLAY "FAIL: want [CC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

