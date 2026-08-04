*> vybe-test: cobol/category_string_overflow/test_str_ovf_multiple
*> origin: languages/cobol/tests/cobol/test_category_string_overflow.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(2) VALUE 'AB'. 01 S2 PIC X(2) VALUE 'CD'. 01 R PIC X(3). PROCEDURE DIVISION. STRING S1 DELIMITED BY SIZE S2 DELIMITED BY SIZE INTO R ON OVERFLOW DISPLAY 'OVF'. STOP RUN.

