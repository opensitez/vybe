*> vybe-test: cobol/category_evaluate/test_evaluate_basic_condition
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN 1 DISPLAY "ONE"
              WHEN 2 DISPLAY "TWO"
              WHEN 3 DISPLAY "THREE"
              WHEN OTHER DISPLAY "OTHER"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE"
        DISPLAY "FAIL: want [ONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

