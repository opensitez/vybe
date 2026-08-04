*> vybe-test: cobol/category_data_division_occurs/test_occurs_key_is_subscripted
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 TBL. 05 EL OCCURS 3 TIMES ASCENDING KEY K1. 10 G1 OCCURS 2 TIMES. 15 K1 PIC 9 VALUE 1. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.

