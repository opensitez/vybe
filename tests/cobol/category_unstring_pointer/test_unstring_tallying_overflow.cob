*> vybe-test: cobol/category_unstring_pointer/test_unstring_tallying_overflow
*> origin: languages/cobol/tests/cobol/test_category_unstring_pointer.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X(1). 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 TALLYING IN T ON OVERFLOW DISPLAY 'OVF'. STOP RUN.

