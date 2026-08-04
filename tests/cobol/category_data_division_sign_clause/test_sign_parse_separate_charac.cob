*> vybe-test: cobol/category_data_division_sign_clause/test_sign_parse_separate_character_runtime
*> origin: languages/cobol/tests/cobol/test_category_data_division_sign_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC S9(3) VALUE -12 IS EXTERNAL SIGN TRAILING SEPARATE CHARACTER. PROCEDURE DIVISION. DISPLAY V.
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12-"
        DISPLAY "FAIL: want [12-] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

