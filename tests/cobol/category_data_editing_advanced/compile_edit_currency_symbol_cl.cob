*> vybe-test: cobol/category_data_editing_advanced/compile_edit_currency_symbol_clause
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC $$,999.
01 Y PIC 999 VALUE 12.
PROCEDURE DIVISION.
    MOVE Y TO X
    STOP RUN.

