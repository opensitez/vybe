*> vybe-test: cobol/supported_types_expanded/pointer_usage_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR USAGE IS POINTER.
PROCEDURE DIVISION.
    SET WS-PTR TO NULL.
    STOP RUN.

