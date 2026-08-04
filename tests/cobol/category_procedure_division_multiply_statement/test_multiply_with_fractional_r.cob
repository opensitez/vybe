*> vybe-test: cobol/category_procedure_division_multiply_statement/test_multiply_with_fractional_result_runtime
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_multiply_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9V9 VALUE 2.5. 01 V2 PIC 9 VALUE 2. PROCEDURE DIVISION. MULTIPLY V2 BY V1 GIVING V1. DISPLAY V1 STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING V1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "05"
        DISPLAY "FAIL: want [05] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

