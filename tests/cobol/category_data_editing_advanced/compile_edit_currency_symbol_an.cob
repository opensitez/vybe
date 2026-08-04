*> vybe-test: cobol/category_data_editing_advanced/compile_edit_currency_symbol_and_comma
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC $$,$$$.99.
01 Y PIC 9999V99 VALUE 123456.
PROCEDURE DIVISION.
    MOVE Y TO X.
    STOP RUN.

