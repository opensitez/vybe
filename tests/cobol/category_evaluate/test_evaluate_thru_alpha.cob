*> vybe-test: cobol/category_evaluate/test_evaluate_thru_alpha
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-THRU-ALPHA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC X VALUE "M".
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN "A" THRU "H" DISPLAY "GROUP 1"
              WHEN "I" THRU "P" DISPLAY "GROUP 2"
              WHEN "Q" THRU "Z" DISPLAY "GROUP 3"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "GROUP 1" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "GROUP 2"
        DISPLAY "FAIL: want [GROUP 2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

