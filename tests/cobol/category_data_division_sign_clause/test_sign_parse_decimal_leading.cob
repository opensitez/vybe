*> vybe-test: cobol/category_data_division_sign_clause/test_sign_parse_decimal_leading
*> origin: languages/cobol/tests/cobol/test_category_data_division_sign_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC S9(3)V99 SIGN IS LEADING VALUE -12.34. PROCEDURE DIVISION. IF V IS NEGATIVE DISPLAY 'NEG' ELSE DISPLAY 'POS' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'NEG' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NEG"
        DISPLAY "FAIL: want [NEG] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

