*> vybe-test: cobol/category_evaluate/test_evaluate_also_any
*> origin: languages/cobol/tests/cobol/test_category_evaluate.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-ANY-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL-1 PIC X VALUE "C".
       01 VAL-2 PIC 9 VALUE 5.
       PROCEDURE DIVISION.
           EVALUATE VAL-1 ALSO VAL-2
              WHEN "A" ALSO ANY DISPLAY "A-ANY"
              WHEN ANY ALSO 5 DISPLAY "ANY-5"
           END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A-ANY" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ANY-5"
        DISPLAY "FAIL: want [ANY-5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

