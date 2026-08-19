*> vybe-test: cobol/category_intrinsic_function_upper_case/test_upper_case_special
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_upper_case.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) VALUE 'h@ll!'. PROCEDURE DIVISION. DISPLAY FUNCTION UPPER-CASE(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION UPPER-CASE(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "H@LL!"
        DISPLAY "FAIL: want [H@LL!] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

