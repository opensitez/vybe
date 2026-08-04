*> vybe-test: cobol/category_procedure_division_multiply_statement/test_multiply_rounded
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_multiply_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9V9 VALUE 2.5. 01 V2 PIC 9V9 VALUE 3.0. 01 R PIC 9. PROCEDURE DIVISION. MULTIPLY V1 BY V2 GIVING R ROUNDED. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "8"
        DISPLAY "FAIL: want [8] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

