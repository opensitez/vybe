*> vybe-test: cobol/category_procedure_division_subtract_statement/test_subtract_size_error_false_path
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_subtract_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 9. 01 V2 PIC 9 VALUE 1. 01 R PIC 9. PROCEDURE DIVISION. SUBTRACT V2 FROM V1 ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK' END-SUBTRACT DISPLAY R STOP RUN.

