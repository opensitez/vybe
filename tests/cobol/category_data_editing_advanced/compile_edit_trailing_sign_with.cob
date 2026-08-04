*> vybe-test: cobol/category_data_editing_advanced/compile_edit_trailing_sign_with_currency
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC ZZ9CR.
01 Y PIC S999 VALUE -12.
PROCEDURE DIVISION.
    MOVE Y TO X.
    STOP RUN.

