*> vybe-test: cobol/category_unstring_overflow/test_unstr_overflow_trigger
*> origin: languages/cobol/tests/cobol/test_category_unstring_overflow.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S PIC X(5) VALUE 'A*B*C'. 01 R1 PIC X. 01 P PIC 9 VALUE 6. PROCEDURE DIVISION. UNSTRING S DELIMITED BY '*' INTO R1 WITH POINTER P ON OVERFLOW DISPLAY 'OVF'. STOP RUN.

