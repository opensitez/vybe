*> vybe-test: cobol/value_all_forms/value_low_value_singular_compiles
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(4) VALUE LOW-VALUE.
PROCEDURE DIVISION.
    DISPLAY S.
    STOP RUN.

