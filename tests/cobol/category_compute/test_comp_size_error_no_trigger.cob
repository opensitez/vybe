*> vybe-test: cobol/category_compute/test_comp_size_error_no_trigger
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC 9. PROCEDURE DIVISION. COMPUTE R = 9 ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK'. STOP RUN.

