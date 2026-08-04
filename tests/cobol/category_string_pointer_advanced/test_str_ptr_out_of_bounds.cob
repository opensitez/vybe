*> vybe-test: cobol/category_string_pointer_advanced/test_str_ptr_out_of_bounds
*> origin: languages/cobol/tests/cobol/test_category_string_pointer_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(2) VALUE 'AB'. 01 R PIC X(4). 01 P PIC 9 VALUE 5. PROCEDURE DIVISION. STRING S1 DELIMITED BY SIZE INTO R WITH POINTER P ON OVERFLOW DISPLAY 'OVF'. STOP RUN.

