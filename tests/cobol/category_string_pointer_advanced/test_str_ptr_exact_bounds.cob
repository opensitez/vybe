*> vybe-test: cobol/category_string_pointer_advanced/test_str_ptr_exact_bounds
*> origin: languages/cobol/tests/cobol/test_category_string_pointer_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X VALUE 'A'. 01 R PIC X(4). 01 P PIC 9 VALUE 4. PROCEDURE DIVISION. STRING S1 DELIMITED BY SIZE INTO R WITH POINTER P. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "   A"
        DISPLAY "FAIL: want [   A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

