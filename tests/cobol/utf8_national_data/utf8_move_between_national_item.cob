*> vybe-test: cobol/utf8_national_data/utf8_move_between_national_items_compiles
*> origin: languages/cobol/tests/cobol/test_utf8_national_data.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. UTF5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
01 N2 PIC N(10).
PROCEDURE DIVISION.
    MOVE N1 TO N2.
    STOP RUN.

