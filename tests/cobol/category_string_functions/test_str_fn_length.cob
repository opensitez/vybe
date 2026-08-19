*> vybe-test: cobol/category_string_functions/test_str_fn_length
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE 'ABC'. PROCEDURE DIVISION. DISPLAY FUNCTION LENGTH(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION LENGTH(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

