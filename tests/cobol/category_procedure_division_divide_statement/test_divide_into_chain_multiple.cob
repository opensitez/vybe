*> vybe-test: cobol/category_procedure_division_divide_statement/test_divide_into_chain_multiple_quotients
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_divide_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 DIVIDEND PIC 9 VALUE 9. 01 DIVISOR PIC 9 VALUE 2. PROCEDURE DIVISION. DIVIDE DIVIDEND INTO DIVISOR. DISPLAY DIVISOR.
    MOVE SPACES TO WS-VYBE-L
    STRING DIVISOR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

