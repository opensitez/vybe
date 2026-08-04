*> vybe-test: cobol/initialize_forms/initialize_alphanumeric_category_replacing
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "HELLO".
PROCEDURE DIVISION.
    INITIALIZE S REPLACING ALPHANUMERIC DATA BY "_".
    STOP RUN.

