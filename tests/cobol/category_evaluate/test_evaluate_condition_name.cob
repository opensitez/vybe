*> vybe-test: cobol/category_evaluate/test_evaluate_condition_name
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-COND-NAME.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STATUS-CODE PIC X VALUE "P".
          88 IS-ACTIVE VALUE "A".
          88 IS-PENDING VALUE "P".
       PROCEDURE DIVISION.
           EVALUATE TRUE
              WHEN IS-ACTIVE DISPLAY "ACTIVE"
              WHEN IS-PENDING DISPLAY "PENDING"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ACTIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ACTIVE"
        DISPLAY "FAIL: want [ACTIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

