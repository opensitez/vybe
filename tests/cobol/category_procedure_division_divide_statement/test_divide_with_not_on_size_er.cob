*> vybe-test: cobol/category_procedure_division_divide_statement/test_divide_with_not_on_size_error_path
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_divide_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 8. 01 V2 PIC 9 VALUE 2. 01 R PIC 9. 01 REM PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R REMAINDER REM ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY R DISPLAY REM END-DIVIDE STOP RUN.

