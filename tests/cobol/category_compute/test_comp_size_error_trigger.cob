*> vybe-test: cobol/category_compute/test_comp_size_error_trigger
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC 9. PROCEDURE DIVISION. COMPUTE R = 10 ON SIZE ERROR DISPLAY 'ERR'. STOP RUN.

