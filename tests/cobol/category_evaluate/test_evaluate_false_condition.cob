*> vybe-test: cobol/category_evaluate/test_evaluate_false_condition
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-FALSE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 FLAG PIC X VALUE "Y".
       PROCEDURE DIVISION.
           EVALUATE FALSE
              WHEN FLAG = "Y" DISPLAY "FLAG IS NOT Y"
              WHEN FLAG = "N" DISPLAY "FLAG IS NOT N"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "FLAG IS NOT Y" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FLAG IS NOT N"
        DISPLAY "FAIL: want [FLAG IS NOT N] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

