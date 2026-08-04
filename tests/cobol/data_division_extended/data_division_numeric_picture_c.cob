*> vybe-test: cobol/data_division_extended/data_division_numeric_picture_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-N PIC 9(5) VALUE 12345.
PROCEDURE DIVISION.
    DISPLAY WS-N.
    STOP RUN.

