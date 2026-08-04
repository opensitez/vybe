*> vybe-test: cobol/category_procedure_division_multiply_statement/test_multiply_by_literal
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_multiply_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 3. 01 V2 PIC 99 VALUE 0. PROCEDURE DIVISION. MULTIPLY 11 BY V1 GIVING V2 DISPLAY V2 STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING V2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "33"
        DISPLAY "FAIL: want [33] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

