*> vybe-test: cobol/data_division_extended/data_division_alphanumeric_picture_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-S PIC X(8) VALUE "HELLO".
PROCEDURE DIVISION.
    DISPLAY WS-S.
    STOP RUN.

