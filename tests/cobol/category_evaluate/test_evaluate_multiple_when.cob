*> vybe-test: cobol/category_evaluate/test_evaluate_multiple_when
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC 9 VALUE 4.
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN 1 
              WHEN 3 
              WHEN 5 DISPLAY "ODD"
              WHEN 2
              WHEN 4
              WHEN 6 DISPLAY "EVEN"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ODD" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "EVEN"
        DISPLAY "FAIL: want [EVEN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

