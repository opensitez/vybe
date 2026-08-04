*> vybe-test: cobol/category_pointers/test_pointers_set_null
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-NULL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO NULL.
           IF WS-PTR = NULL
              DISPLAY "IS NULL"
           END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "IS NULL" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "IS NULL"
        DISPLAY "FAIL: want [IS NULL] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

