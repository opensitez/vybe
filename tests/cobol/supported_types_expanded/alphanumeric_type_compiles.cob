*> vybe-test: cobol/supported_types_expanded/alphanumeric_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "HELLO".
PROCEDURE DIVISION.
    DISPLAY WS-TXT.
    STOP RUN.

