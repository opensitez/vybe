*> vybe-test: cobol/category_evaluate/test_evaluate_also_basic
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-ALSO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL-1 PIC X VALUE "A".
       01 VAL-2 PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           EVALUATE VAL-1 ALSO VAL-2
              WHEN "A" ALSO 1 DISPLAY "A1"
              WHEN "A" ALSO 2 DISPLAY "A2"
              WHEN "B" ALSO 2 DISPLAY "B2"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A1" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A1"
        DISPLAY "FAIL: want [A1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

