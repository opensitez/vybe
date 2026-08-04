*> vybe-test: cobol/data_division_extended/data_division_group_item_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRP.
   05 WS-A PIC X(2) VALUE "AB".
   05 WS-B PIC X(2) VALUE "CD".
PROCEDURE DIVISION.
    DISPLAY WS-GRP.
    STOP RUN.

