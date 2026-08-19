*> vybe-test: cobol/category_data_division_justified_right/test_justified_parse_7
*> origin: languages/cobol/tests/cobol/test_category_data_division_justified_right.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(6) JUSTIFIED RIGHT. 01 W PIC X(6) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'A' TO W MOVE 'ABCD' TO V DISPLAY '[' V ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE V DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[  ABCD]"
        DISPLAY "FAIL: want [[  ABCD]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. DISPLAY '[' W ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE W DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[     A]"
        DISPLAY "FAIL: want [[     A]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

