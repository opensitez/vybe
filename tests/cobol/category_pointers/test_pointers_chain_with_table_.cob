*> vybe-test: cobol/category_pointers/test_pointers_chain_with_table_entry
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-TBL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-BUF.
           05 ENTRIES OCCURS 2 TIMES PIC X(4) VALUE 'AAAA'.
       01 WS-PTR USAGE POINTER.
       01 WS-VIEW PIC X(4).
       PROCEDURE DIVISION.
           SET WS-PTR TO ADDRESS OF ENTRIES(2).
           SET ADDRESS OF WS-VIEW TO WS-PTR.
           DISPLAY WS-VIEW.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-VIEW DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAAA"
        DISPLAY "FAIL: want [AAAA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

