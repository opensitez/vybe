*> vybe-test: cobol/utf8_national_data/utf8_compare_national_literal_compiles
*> origin: languages/cobol/tests/cobol/test_utf8_national_data.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. UTF7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
PROCEDURE DIVISION.
    IF N1 = N"CAFE" DISPLAY "Y" END-IF.
    STOP RUN.

