*> vybe-test: cobol/category_procedure_division_multiply_statement/test_multiply_not_size_error_path
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_multiply_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 6. 01 V2 PIC 9 VALUE 7. 01 R PIC 99 VALUE 0. PROCEDURE DIVISION. MULTIPLY V1 BY V2 GIVING R ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK' END-MULTIPLY DISPLAY R STOP RUN.

