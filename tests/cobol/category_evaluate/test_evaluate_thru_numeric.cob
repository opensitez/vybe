*> vybe-test: cobol/category_evaluate/test_evaluate_thru_numeric
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-THRU-NUM.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC 99 VALUE 15.
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN 01 THRU 10 DISPLAY "1-10"
              WHEN 11 THRU 20 DISPLAY "11-20"
              WHEN 21 THRU 30 DISPLAY "21-30"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "1-10" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "11-20"
        DISPLAY "FAIL: want [11-20] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

