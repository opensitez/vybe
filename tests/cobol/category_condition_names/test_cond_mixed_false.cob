*> vybe-test: cobol/category_condition_names/test_cond_mixed_false
*> origin: languages/cobol/tests/cobol/test_category_condition_names.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 9 VALUE 4. 88 VALID VALUE 1 3 5 THRU 9. PROCEDURE DIVISION. IF VALID DISPLAY 'Y' ELSE DISPLAY 'N' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "N"
        DISPLAY "FAIL: want [N] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

