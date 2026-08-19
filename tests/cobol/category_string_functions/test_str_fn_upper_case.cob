*> vybe-test: cobol/category_string_functions/test_str_fn_upper_case
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(3) VALUE 'abc'. PROCEDURE DIVISION. DISPLAY FUNCTION UPPER-CASE(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION UPPER-CASE(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC"
        DISPLAY "FAIL: want [ABC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

