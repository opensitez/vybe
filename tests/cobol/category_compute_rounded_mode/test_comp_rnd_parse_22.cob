*> vybe-test: cobol/category_compute_rounded_mode/test_comp_rnd_parse_22
*> origin: languages/cobol/tests/cobol/test_category_compute_rounded_mode.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC 99 VALUE 0. PROCEDURE DIVISION. COMPUTE R ROUNDED MODE IS PROHIBITED = 8.5 ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY 'OK' END-COMPUTE.

