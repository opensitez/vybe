*> vybe-test: cobol/category_compiler_directives/test_eject_skip_directive
*> origin: languages/cobol/tests/cobol/test_category_compiler_directives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PAGE-DIR.
       EJECT.
       SKIP1.
       SKIP2.
       SKIP3.
       PROCEDURE DIVISION.
           DISPLAY "PAGING PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "PAGING PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "PAGING PARSED"
        DISPLAY "FAIL: want [PAGING PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

