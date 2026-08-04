*> vybe-test: cobol/category_procedure_division_multiply_statement/test_multiply_size_error
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_multiply_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 5. 01 V2 PIC 9 VALUE 3. 01 R PIC 9. PROCEDURE DIVISION. MULTIPLY V1 BY V2 GIVING R ON SIZE ERROR DISPLAY 'ERR'. STOP RUN.

