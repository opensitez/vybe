*> vybe-test: cobol/category_string_functions/test_str_fn_trim_leading
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE '  ABC'. PROCEDURE DIVISION. DISPLAY '[' FUNCTION TRIM(V LEADING) ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE FUNCTION TRIM(V LEADING) DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[ABC]"
        DISPLAY "FAIL: want [[ABC]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

