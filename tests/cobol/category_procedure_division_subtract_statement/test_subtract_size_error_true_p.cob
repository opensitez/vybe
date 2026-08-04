*> vybe-test: cobol/category_procedure_division_subtract_statement/test_subtract_size_error_true_path
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_subtract_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 0. 01 V2 PIC 9 VALUE 1. 01 V PIC 99 VALUE 1. PROCEDURE DIVISION. SUBTRACT V1 FROM V GIVING V2 ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK' END-SUBTRACT DISPLAY V2 STOP RUN.

