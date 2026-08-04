*> vybe-test: cobol/category_evaluate_complex/test_eval_thru_first
*> origin: languages/cobol/tests/cobol/test_category_evaluate_complex.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 9 VALUE 3. 01 Y PIC 9 VALUE 5. PROCEDURE DIVISION. EVALUATE X ALSO Y WHEN 1 THRU 4 ALSO 5 DISPLAY 'Y' END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

