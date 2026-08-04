*> vybe-test: cobol/initialize_forms/initialize_replacing_numeric_by_literal
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3).
01 B PIC 9(3).
PROCEDURE DIVISION.
    INITIALIZE A B REPLACING NUMERIC DATA BY 5.
    STOP RUN.

