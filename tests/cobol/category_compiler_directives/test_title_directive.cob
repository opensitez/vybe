*> vybe-test: cobol/category_compiler_directives/test_title_directive
*> origin: languages/cobol/tests/cobol/test_category_compiler_directives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TITLE-DIR.
       TITLE "MY TITLE".
       PROCEDURE DIVISION.
           DISPLAY "TITLE PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "TITLE PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "TITLE PARSED"
        DISPLAY "FAIL: want [TITLE PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

