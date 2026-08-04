*> vybe-test: cobol/supported_types_expanded/binary_usage_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BIN PIC 9(4) USAGE IS BINARY VALUE 7.
PROCEDURE DIVISION.
    ADD 1 TO WS-BIN.
    STOP RUN.

