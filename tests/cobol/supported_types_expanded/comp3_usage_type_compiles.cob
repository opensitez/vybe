*> vybe-test: cobol/supported_types_expanded/comp3_usage_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PACK PIC 9(5) USAGE IS COMP-3 VALUE 20.
PROCEDURE DIVISION.
    ADD 5 TO WS-PACK.
    STOP RUN.

