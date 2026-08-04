*> vybe-test: cobol/category_evaluate/test_evaluate_true_condition
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-TRUE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL-A PIC 9 VALUE 5.
       01 VAL-B PIC 9 VALUE 10.
       PROCEDURE DIVISION.
           EVALUATE TRUE
              WHEN VAL-A > 10 DISPLAY "A>10"
              WHEN VAL-B > 5  DISPLAY "B>5"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A>10" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B>5"
        DISPLAY "FAIL: want [B>5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

