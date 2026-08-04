*> vybe-test: cobol/initialize_forms/initialize_alphabetic_category
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC A(5) VALUE "HELLO".
PROCEDURE DIVISION.
    INITIALIZE S REPLACING ALPHABETIC DATA BY "X".
    STOP RUN.

