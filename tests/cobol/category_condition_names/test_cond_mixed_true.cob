*> vybe-test: cobol/category_condition_names/test_cond_mixed_true
*> origin: languages/cobol/tests/cobol/test_category_condition_names.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 9 VALUE 8. 88 VALID VALUE 1 3 5 THRU 9. PROCEDURE DIVISION. IF VALID DISPLAY 'Y' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

