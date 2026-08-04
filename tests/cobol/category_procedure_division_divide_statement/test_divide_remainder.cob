*> vybe-test: cobol/category_procedure_division_divide_statement/test_divide_remainder
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_divide_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V1 PIC 9 VALUE 7. 01 V2 PIC 9 VALUE 2. 01 R PIC 9. 01 REM PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R REMAINDER REM. DISPLAY R REM.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE REM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "31"
        DISPLAY "FAIL: want [31] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

