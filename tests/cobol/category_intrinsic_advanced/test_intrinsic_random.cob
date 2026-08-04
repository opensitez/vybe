*> vybe-test: cobol/category_intrinsic_advanced/test_intrinsic_random
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-RND.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 RES PIC V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION RANDOM(123).
           DISPLAY "RANDOM PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "RANDOM PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "RANDOM PARSED"
        DISPLAY "FAIL: want [RANDOM PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

