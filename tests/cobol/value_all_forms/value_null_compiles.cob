*> vybe-test: cobol/value_all_forms/value_null_compiles
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER VALUE NULL.
PROCEDURE DIVISION.
    CONTINUE.
    STOP RUN.

