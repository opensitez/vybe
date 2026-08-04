*> vybe-test: cobol/category_pointers/test_pointers_null_after_set
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-NULL-2.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO NULL.
           IF WS-PTR = NULL
              DISPLAY "NULL".
    MOVE SPACES TO WS-VYBE-L
    STRING "NULL" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NULL"
        DISPLAY "FAIL: want [NULL] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           END-IF.
           STOP RUN.

