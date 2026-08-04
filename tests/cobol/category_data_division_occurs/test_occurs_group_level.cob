*> vybe-test: cobol/category_data_division_occurs/test_occurs_group_level
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 TBL OCCURS 3 TIMES. 05 EL PIC 9 VALUE 1. PROCEDURE DIVISION. DISPLAY EL(1). STOP RUN.

