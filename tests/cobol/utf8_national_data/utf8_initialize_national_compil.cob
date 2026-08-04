*> vybe-test: cobol/utf8_national_data/utf8_initialize_national_compiles
*> origin: languages/cobol/tests/cobol/test_utf8_national_data.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. UTF6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
PROCEDURE DIVISION.
    INITIALIZE N1.
    STOP RUN.

