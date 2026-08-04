*> vybe-test: cobol/initialize_forms/initialize_numeric_category
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(5) VALUE 12345.
PROCEDURE DIVISION.
    INITIALIZE N REPLACING NUMERIC DATA BY 9.
    STOP RUN.

