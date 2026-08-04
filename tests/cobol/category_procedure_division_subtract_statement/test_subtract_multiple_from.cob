*> vybe-test: cobol/category_procedure_division_subtract_statement/test_subtract_multiple_from
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_subtract_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 5. 01 V2 PIC 9 VALUE 6. PROCEDURE DIVISION. SUBTRACT 2 FROM V1 V2. DISPLAY V1 V2.
    MOVE SPACES TO WS-VYBE-L
    STRING V1 DELIMITED SIZE V2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "34"
        DISPLAY "FAIL: want [34] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

