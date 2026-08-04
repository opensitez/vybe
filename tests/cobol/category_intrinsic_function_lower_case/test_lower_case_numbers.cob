*> vybe-test: cobol/category_intrinsic_function_lower_case/test_lower_case_numbers
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_lower_case.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE 'H3LL0'. PROCEDURE DIVISION. DISPLAY FUNCTION LOWER-CASE(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE LOWER-CASE(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "h3ll0"
        DISPLAY "FAIL: want [h3ll0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

