*> vybe-test: cobol/category_string_delimiters/test_str_overflow_no_trigger
*> origin: languages/cobol/tests/cobol/test_category_string_delimiters.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(2) VALUE 'AB'. 01 R PIC X(5). PROCEDURE DIVISION. STRING S1 DELIMITED BY SIZE INTO R ON OVERFLOW DISPLAY 'OVF' NOT ON OVERFLOW DISPLAY 'OK'. STOP RUN.

