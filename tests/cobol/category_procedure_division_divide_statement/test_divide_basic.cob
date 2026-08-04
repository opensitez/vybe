*> vybe-test: cobol/category_procedure_division_divide_statement/test_divide_basic
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_divide_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 6. 01 V2 PIC 9 VALUE 2. PROCEDURE DIVISION. DIVIDE V2 INTO V1. DISPLAY V1.
    MOVE SPACES TO WS-VYBE-L
    STRING V1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

