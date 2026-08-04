*> vybe-test: cobol/category_procedure_division_divide_statement/test_divide_zero_divide
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_divide_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 5. 01 V2 PIC 9 VALUE 0. 01 R PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R ON SIZE ERROR DISPLAY 'ERR'. STOP RUN.

