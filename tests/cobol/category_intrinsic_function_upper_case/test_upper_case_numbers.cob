*> vybe-test: cobol/category_intrinsic_function_upper_case/test_upper_case_numbers
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_upper_case.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE 'h3LL0'. PROCEDURE DIVISION. DISPLAY FUNCTION UPPER-CASE(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION UPPER-CASE(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "H3LL0"
        DISPLAY "FAIL: want [H3LL0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

