*> vybe-test: cobol/supported_types_expanded/function_pointer_usage_type_compiles
*> origin: languages/cobol/tests/cobol/test_supported_types_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FPTR USAGE IS FUNCTION-POINTER.
PROCEDURE DIVISION.
    DISPLAY "FPTR".
    STOP RUN.

