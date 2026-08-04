*> vybe-test: cobol/category_data_division_advanced/test_dd_sign_clause_local_group
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC S9(2) VALUE -12 SIGN IS LEADING. PROCEDURE DIVISION. DISPLAY V.
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-12"
        DISPLAY "FAIL: want [-12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

