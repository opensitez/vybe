*> vybe-test: cobol/category_data_division_global_clause/test_global_parse_parent_modifies_after_call_start
*> origin: languages/cobol/tests/cobol/test_category_data_division_global_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G PIC 9 VALUE 3 IS GLOBAL. PROCEDURE DIVISION. CALL 'CHILD'. ADD 2 TO G. DISPLAY G.
    MOVE SPACES TO WS-VYBE-L
    STRING G DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9"
        DISPLAY "FAIL: want [9] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. CHILD. PROCEDURE DIVISION. ADD 4 TO G. EXIT PROGRAM. END PROGRAM CHILD. END PROGRAM T.

