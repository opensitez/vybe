*> vybe-test: cobol/category_procedure_division_add_statement/test_add_size_error_false_path
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_add_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 99 VALUE 99. 01 B PIC 99 VALUE 1. 01 R PIC 9 VALUE 0. PROCEDURE DIVISION. ADD A TO B ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK' END-ADD STOP RUN.

