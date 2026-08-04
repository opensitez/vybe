*> vybe-test: cobol/category_evaluate/test_evaluate_partial_also
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-PARTIAL-ALSO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL-1 PIC X VALUE "X".
       01 VAL-2 PIC X VALUE "Y".
       PROCEDURE DIVISION.
           EVALUATE VAL-1 ALSO VAL-2
              WHEN "X" ALSO "Z" THRU "Y" DISPLAY "BAD"
              WHEN "X" ALSO "A" THRU "Z" DISPLAY "GOOD"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "BAD" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "GOOD"
        DISPLAY "FAIL: want [GOOD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

