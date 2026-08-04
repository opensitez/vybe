*> vybe-test: cobol/category_data_division_sign_clause/test_sign_parse_zero_negative_check
*> origin: languages/cobol/tests/cobol/test_category_data_division_sign_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC S9(3) VALUE 0 SIGN IS LEADING SEPARATE. PROCEDURE DIVISION. IF V IS ZERO DISPLAY 'ZERO' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'ZERO' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZERO"
        DISPLAY "FAIL: want [ZERO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

