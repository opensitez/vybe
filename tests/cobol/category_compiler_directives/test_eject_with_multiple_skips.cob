*> vybe-test: cobol/category_compiler_directives/test_eject_with_multiple_skips
*> origin: languages/cobol/tests/cobol/test_category_compiler_directives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EJECT-MULTI.
       EJECT.
       SKIP1.
       SKIP3.
       PROCEDURE DIVISION.
           DISPLAY "EJECT-MULTI".
    MOVE SPACES TO WS-VYBE-L
    STRING "EJECT-MULTI" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "EJECT-MULTI"
        DISPLAY "FAIL: want [EJECT-MULTI] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

