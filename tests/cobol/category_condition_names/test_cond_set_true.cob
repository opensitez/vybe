*> vybe-test: cobol/category_condition_names/test_cond_set_true
*> origin: languages/cobol/tests/cobol/test_category_condition_names.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 9. 88 IS-FIVE VALUE 5. PROCEDURE DIVISION. SET IS-FIVE TO TRUE. DISPLAY X.
    MOVE SPACES TO WS-VYBE-L
    STRING X DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

