*> vybe-test: cobol/utf8_national_data/utf8_national_group_compiles
*> origin: languages/cobol/tests/cobol/test_utf8_national_data.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. UTF9.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NG.
   05 N1 PIC N(5).
   05 N2 PIC N(5).
PROCEDURE DIVISION.
    STOP RUN.

