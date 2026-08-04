*> vybe-test: cobol/utf8_national_data/utf8_function_pair_compiles
*> origin: languages/cobol/tests/cobol/test_utf8_national_data.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. UTF10.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
01 D1 PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION NATIONAL-OF(D1) TO N1.
    MOVE FUNCTION DISPLAY-OF(N1) TO D1.
    STOP RUN.

