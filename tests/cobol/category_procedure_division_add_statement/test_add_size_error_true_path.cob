*> vybe-test: cobol/category_procedure_division_add_statement/test_add_size_error_true_path
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_add_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9999 VALUE 9999. 01 B PIC 9 VALUE 9. 01 R PIC 9 VALUE 0. PROCEDURE DIVISION. ADD A TO B ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK' END-ADD DISPLAY R STOP RUN.

