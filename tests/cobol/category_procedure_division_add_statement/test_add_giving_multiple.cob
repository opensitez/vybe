*> vybe-test: cobol/category_procedure_division_add_statement/test_add_giving_multiple
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_add_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 1. 01 V2 PIC 9 VALUE 2. 01 R1 PIC 9. 01 R2 PIC 9. PROCEDURE DIVISION. ADD V1 V2 GIVING R1 R2. DISPLAY R1 R2.
    MOVE SPACES TO WS-VYBE-L
    STRING R1 DELIMITED SIZE R2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "33"
        DISPLAY "FAIL: want [33] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

