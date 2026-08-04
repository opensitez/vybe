*> vybe-test: cobol/category_data_division_justified_right/test_justified_parse_6
*> origin: languages/cobol/tests/cobol/test_category_data_division_justified_right.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(6) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'BC' TO V. DISPLAY '[' V ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE V DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[    BC]"
        DISPLAY "FAIL: want [[    BC]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

