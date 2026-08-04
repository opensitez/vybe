*> vybe-test: cobol/supported_types_expanded/decimal_implied_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AMT PIC 9(3)V99 VALUE 123.45.
PROCEDURE DIVISION.
    DISPLAY WS-AMT.
    STOP RUN.

