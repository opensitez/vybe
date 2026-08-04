*> vybe-test: cobol/data_division_extended/data_division_level_77_item_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-VAL PIC 9(3) VALUE 100.
PROCEDURE DIVISION.
    DISPLAY WS-VAL.
    STOP RUN.

