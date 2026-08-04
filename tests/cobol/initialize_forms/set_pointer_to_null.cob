*> vybe-test: cobol/initialize_forms/set_pointer_to_null
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PTR USAGE POINTER.
PROCEDURE DIVISION.
    SET PTR TO NULL.
    STOP RUN.

