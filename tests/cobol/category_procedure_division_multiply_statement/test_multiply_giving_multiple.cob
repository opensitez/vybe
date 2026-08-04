*> vybe-test: cobol/category_procedure_division_multiply_statement/test_multiply_giving_multiple
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_multiply_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 2. 01 V2 PIC 9 VALUE 3. 01 R1 PIC 99. 01 R2 PIC 99. PROCEDURE DIVISION. MULTIPLY V1 BY V2 GIVING R1 R2. DISPLAY R1 R2.
    MOVE SPACES TO WS-VYBE-L
    STRING R1 DELIMITED SIZE R2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0606"
        DISPLAY "FAIL: want [0606] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

