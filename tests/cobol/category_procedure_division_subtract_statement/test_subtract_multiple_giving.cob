*> vybe-test: cobol/category_procedure_division_subtract_statement/test_subtract_multiple_giving
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_subtract_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 2. 01 V2 PIC 9 VALUE 1. 01 V3 PIC 9 VALUE 7. 01 R PIC 9. PROCEDURE DIVISION. SUBTRACT V1 V2 FROM V3 GIVING R. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "4"
        DISPLAY "FAIL: want [4] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

