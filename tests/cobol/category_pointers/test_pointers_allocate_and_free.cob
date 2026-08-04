*> vybe-test: cobol/category_pointers/test_pointers_allocate_and_free
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-FREE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-PTR USAGE POINTER.
       01 WS-TGT PIC X(4).
       PROCEDURE DIVISION.
           ALLOCATE WS-TGT RETURNING WS-PTR.
           IF WS-PTR = NULL
               DISPLAY "NO"
           ELSE
               DISPLAY "YES"
           END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "NO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "YES"
        DISPLAY "FAIL: want [YES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           FREE WS-PTR.
           STOP RUN.

