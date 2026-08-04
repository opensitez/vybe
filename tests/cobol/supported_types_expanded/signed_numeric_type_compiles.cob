*> vybe-test: cobol/supported_types_expanded/signed_numeric_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC S9(4) VALUE -25.
PROCEDURE DIVISION.
    DISPLAY WS-NUM.
    STOP RUN.

