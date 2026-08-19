*> vybe-test: cobol/category_intrinsic_function_length/test_length_numeric
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_length.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC 9(3). PROCEDURE DIVISION. DISPLAY FUNCTION LENGTH(V).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION LENGTH(V) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

