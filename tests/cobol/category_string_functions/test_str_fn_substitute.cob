*> vybe-test: cobol/category_string_functions/test_str_fn_substitute
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE 'AABAA'. PROCEDURE DIVISION. DISPLAY FUNCTION SUBSTITUTE(V 'A' 'X').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION SUBSTITUTE(V, 'A', 'X') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XXBXX"
        DISPLAY "FAIL: want [XXBXX] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

